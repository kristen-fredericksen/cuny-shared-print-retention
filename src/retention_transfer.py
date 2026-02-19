"""
CUNY Shared Print Monograph Trust - Retention Transfer Script

This script helps transfer retention commitments from one school to another
when a school can no longer retain a book.

Phase 1: Lookup and school selection
- Reads barcodes and school codes from Excel file
- Looks up each item using the specified school's API key
- Queries Network Zone to find all schools holding the title
- Selects replacement school based on rules

Phase 2: Draft email generation
- Generates .eml files for each replacement school
- Includes school-specific Primo VE links

Phase 3: Update leaving school's Alma records
- Sets committed_to_retain to No on item record
- Clears retention_reason on item record
- Removes 583 field from holdings record (if no other retained items remain)

Phase 4: Update taking school's Alma records
- Sets committed_to_retain to Yes on item record
- Sets retention_reason to CUNYSharedPrint on item record
- Adds 583 field to holdings record

Usage:
    python retention_transfer.py barcodes.xlsx [output_dir] [--sandbox]

The Excel file should have two columns:
    - Barcode: The item barcode
    - School Code: The Alma Institution Code (e.g., 01CUNY_QC)

Examples:
    python retention_transfer.py data/barcodes.xlsx
    python retention_transfer.py data/barcodes.xlsx output/emails --sandbox
"""

import os
import sys
import re
import requests
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


# =============================================================================
# CONFIGURATION
# =============================================================================

def get_config(sandbox=False):
    """Load configuration from environment variables.

    Args:
        sandbox: If True, use sandbox API keys instead of production keys.
    """
    if sandbox:
        config = {
            "nz_api_key": os.getenv("ALMA_SANDBOX_NZ_API_KEY"),
            "base_url": os.getenv("ALMA_SANDBOX_API_BASE_URL", "https://api-na.hosted.exlibrisgroup.com"),
            "schools_file": os.getenv("ALMA_SANDBOX_SCHOOLS_FILE", "data/schools_sandbox.csv"),
            "sandbox": True
        }
        if not config["nz_api_key"]:
            print("ERROR: No sandbox Network Zone API key found!")
            print("Make sure your .env file has ALMA_SANDBOX_NZ_API_KEY=your_key_here")
            sys.exit(1)
    else:
        config = {
            "nz_api_key": os.getenv("ALMA_NZ_API_KEY"),
            "base_url": os.getenv("ALMA_API_BASE_URL", "https://api-na.hosted.exlibrisgroup.com"),
            "schools_file": os.getenv("SCHOOLS_FILE", "data/schools_template.csv"),
            "sandbox": False
        }
        if not config["nz_api_key"]:
            print("ERROR: No Network Zone API key found!")
            print("Make sure you have a .env file with ALMA_NZ_API_KEY=your_key_here")
            sys.exit(1)

    return config


# =============================================================================
# SCHOOL DATA
# =============================================================================

def load_schools(file_path):
    """
    Load the schools data from CSV file.

    Returns a dictionary keyed by Alma Institution Code.
    """
    try:
        df = pd.read_csv(file_path)
    except FileNotFoundError:
        print(f"ERROR: Schools file not found: {file_path}")
        print("Please create this file with columns: Name, Size, Shared Print, Alma Institution Code, etc.")
        sys.exit(1)

    schools = {}
    for _, row in df.iterrows():
        code = row["Alma Institution Code"]
        schools[code] = {
            "name": row["Name"],
            "size": row["Size"],  # 1 = largest
            "shared_print": row["Shared Print"].strip().lower() == "yes",
            "code": code,
            "oclc_symbol": row.get("OCLC Symbol", ""),
            "primo_view_id": row.get("Primo View ID", ""),
            "chief_librarian_name": row.get("Chief Librarian Name", ""),
            "chief_librarian_email": row.get("Chief Librarian Email", ""),
            "marc_org_code": row.get("MARC Org Code", ""),
            "api_key": row.get("Alma API Key", "")
        }

    return schools


def get_grad_center_code(schools):
    """Find the Grad Center's institution code."""
    for code, school in schools.items():
        if "grad" in school["name"].lower() and "center" in school["name"].lower():
            return code
    return None


# =============================================================================
# ALMA API FUNCTIONS
# =============================================================================

def lookup_item_by_barcode(barcode, schools, base_url):
    """
    Look up an item in Alma by its barcode.

    Since barcode lookup requires an Institution Zone key, we try each
    school's API key until we find the item.

    Returns the item data including which institution holds it.
    """
    url = f"{base_url}/almaws/v1/items"
    headers = {"Accept": "application/json"}

    # Try each school's API key
    for code, school in schools.items():
        api_key = school.get("api_key", "")
        if not api_key or pd.isna(api_key):
            continue

        params = {
            "item_barcode": barcode,
            "apikey": api_key
        }

        try:
            response = requests.get(url, params=params, headers=headers)

            if response.status_code == 200:
                return response.json()
            # 400 means not found at this institution, try next

        except requests.RequestException:
            continue

    return None  # Not found at any institution


def get_nz_mms_id_from_item(item_data):
    """
    Extract the Network Zone MMS ID from item data.

    The NZ MMS ID is stored in the network_number field with format:
    (EXLNZ-01CUNY_NETWORK)991025135669706121
    """
    network_numbers = item_data.get("bib_data", {}).get("network_number", [])

    for nn in network_numbers:
        if "01CUNY_NETWORK" in nn and nn.startswith("(EXLNZ"):
            # Extract the ID after the closing parenthesis
            nz_mms_id = nn.split(")")[-1]
            return nz_mms_id

    return None


def get_holding_institutions_from_nz(nz_mms_id, nz_api_key, base_url):
    """
    Get all institutions that have holdings for a bib record in the Network Zone.

    Parses the MARC AVA fields from the bib record to find holding institutions.

    Returns a list of institution codes (e.g., ['01CUNY_BC', '01CUNY_QC']).
    """
    url = f"{base_url}/almaws/v1/bibs/{nz_mms_id}"

    params = {
        "apikey": nz_api_key,
        "expand": "p_avail"  # Include physical availability info
    }

    headers = {"Accept": "application/json"}

    try:
        response = requests.get(url, params=params, headers=headers)

        if response.status_code != 200:
            print(f"  Could not get NZ bib record (status {response.status_code})")
            return []

        data = response.json()

        # Get the MARC XML from the 'anies' field
        marc_xml = data.get("anies", [""])[0]

        if not marc_xml:
            return []

        # Find all AVA fields (availability info) - these contain institution codes
        # AVA subfield 'a' contains the institution code
        ava_institutions = re.findall(
            r'tag="AVA".*?<subfield code="a">(.*?)</subfield>',
            marc_xml
        )

        # Deduplicate and return
        return list(set(ava_institutions))

    except requests.RequestException as e:
        print(f"  Connection error: {e}")
        return []


def lookup_item_by_barcode(barcode, school_code, schools, base_url):
    """
    Look up an item in Alma by its barcode using a specific school's API key.

    Args:
        barcode: The item barcode to look up
        school_code: The Alma Institution Code (e.g., 01CUNY_QC)
        schools: Dictionary of school data
        base_url: Alma API base URL

    Returns the item data, or None if not found.
    """
    url = f"{base_url}/almaws/v1/items"
    headers = {"Accept": "application/json"}

    # Get the API key for the specified school
    school = schools.get(school_code)
    if not school:
        print(f"  ERROR: School code '{school_code}' not found in schools file")
        return None

    api_key = school.get("api_key", "")
    if not api_key or pd.isna(api_key):
        print(f"  ERROR: No API key configured for {school['name']}")
        return None

    params = {
        "item_barcode": barcode,
        "apikey": api_key
    }

    try:
        response = requests.get(url, params=params, headers=headers)

        if response.status_code == 200:
            return response.json()
        elif response.status_code == 400:
            print(f"  Barcode not found at {school['name']}")
            return None
        else:
            print(f"  API error (status {response.status_code})")
            return None

    except requests.RequestException as e:
        print(f"  Connection error: {e}")
        return None


def check_item_status(item_data):
    """
    Check if the item is eligible for retention transfer based on its status.

    Args:
        item_data: The full item data from Alma API

    Returns:
        (is_eligible, status_description)
        - is_eligible: True if base_status = 1 (Item in place)
        - status_description: Human-readable status
    """
    item_info = item_data.get("item_data", {})
    base_status = item_info.get("base_status", {})

    status_value = base_status.get("value", "")
    status_desc = base_status.get("desc", "Unknown status")

    # base_status = 1 means "Item in place" (eligible)
    # base_status = 0 means "Item not in place" (not eligible)
    is_eligible = status_value == "1"

    return is_eligible, status_desc


def find_holding_institutions(barcode, leaving_school, schools, config):
    """
    Find all institutions that hold a copy of this item.

    1. Look up the barcode using the specified school's API key
    2. Check if item is eligible (base_status = 1)
    3. Extract the Network Zone MMS ID
    4. Query the Network Zone to find all holding institutions

    Args:
        barcode: The item barcode
        leaving_school: The Alma Institution Code of the school leaving retention
        schools: Dictionary of school data
        config: Configuration dictionary

    Returns: (list of institution codes, bib_info dict)
    Returns ("ineligible", bib_info) if item status is not "Item in place"
    """
    # Step 1: Look up the item by barcode using the leaving school's API key
    item_data = lookup_item_by_barcode(barcode, leaving_school, schools, config["base_url"])

    if not item_data:
        return None, None

    # Get basic info
    try:
        iz_mms_id = item_data["bib_data"]["mms_id"]
        title = item_data["bib_data"].get("title", "Unknown title")
    except KeyError:
        print("  Could not extract MMS ID from item data")
        return None, None

    # Step 2: Check if item is eligible based on base_status
    is_eligible, status_desc = check_item_status(item_data)
    if not is_eligible:
        print(f"  ⚠️  Item not eligible: {status_desc}")
        return "ineligible", {
            "mms_id": iz_mms_id,
            "nz_mms_id": None,
            "title": title,
            "item_data": item_data,
            "status": status_desc
        }

    # Step 2: Get the Network Zone MMS ID
    nz_mms_id = get_nz_mms_id_from_item(item_data)

    if not nz_mms_id:
        print("  Could not find Network Zone MMS ID")
        # Fall back to just the originating institution
        return [], {
            "mms_id": iz_mms_id,
            "nz_mms_id": None,
            "title": title,
            "item_data": item_data
        }

    # Step 3: Query Network Zone for all holding institutions
    institutions = get_holding_institutions_from_nz(
        nz_mms_id, config["nz_api_key"], config["base_url"]
    )

    bib_info = {
        "mms_id": iz_mms_id,
        "nz_mms_id": nz_mms_id,
        "title": title,
        "item_data": item_data
    }

    return institutions, bib_info


# =============================================================================
# SCHOOL SELECTION LOGIC
# =============================================================================

def select_replacement_school(holding_institutions, leaving_school, schools):
    """
    Select a replacement school based on the rules:
    1. Exclude schools not in Shared Print
    2. Exclude the leaving school
    3. Pick Grad Center if they hold it
    4. Otherwise pick the largest school that holds it
    5. If no eligible schools, return None (flag for withdrawal review)

    Returns: (selected_school_code, list_of_all_eligible_schools_in_priority_order)
    """
    # Filter to only Shared Print participants who hold the item
    eligible = []
    for inst_code in holding_institutions:
        if inst_code == leaving_school:
            continue  # Skip the leaving school

        if inst_code not in schools:
            continue  # School not in our list

        school = schools[inst_code]
        if not school["shared_print"]:
            continue  # Not a Shared Print participant

        eligible.append(school)

    if not eligible:
        return None, []

    # Check if Grad Center is in the eligible list
    grad_center_code = get_grad_center_code(schools)
    for school in eligible:
        if school["code"] == grad_center_code:
            # Put Grad Center first, then sort rest by size
            others = [s for s in eligible if s["code"] != grad_center_code]
            others.sort(key=lambda s: s["size"])  # 1 = largest, so ascending
            priority_list = [school] + others
            return school["code"], [s["code"] for s in priority_list]

    # No Grad Center - sort by size (1 = largest)
    eligible.sort(key=lambda s: s["size"])

    return eligible[0]["code"], [s["code"] for s in eligible]


# =============================================================================
# PHASE 2: EMAIL GENERATION
# =============================================================================

def build_primo_link(nz_mms_id, school):
    """
    Build a Primo VE link for a Network Zone MMS ID, using the school's Primo instance.

    CUNY's Primo VE URL format:
    https://cuny-{inst_suffix}.primo.exlibrisgroup.com/permalink/{inst_code}/{view_id}/alma{nz_mms_id}

    Args:
        nz_mms_id: Network Zone MMS ID
        school: School dictionary with 'code' and 'primo_view_id'

    Returns:
        Primo VE permalink URL, or None if missing data
    """
    if not nz_mms_id:
        return None

    inst_code = school.get("code", "")
    primo_view_id = school.get("primo_view_id", "")

    if not inst_code or not primo_view_id:
        return None

    # Extract the suffix from institution code (e.g., "01CUNY_BC" -> "bc")
    inst_suffix = inst_code.split("_")[-1].lower()

    return f"https://cuny-{inst_suffix}.primo.exlibrisgroup.com/permalink/{inst_code}/{primo_view_id}/alma{nz_mms_id}"


def generate_draft_email(titles_for_school, replacement_code, schools):
    """
    Generate a draft email for one or more titles going to the same replacement school.

    Args:
        titles_for_school: List of result dictionaries for this school
        replacement_code: Alma Institution Code of the replacement school
        schools: Dictionary of school data

    Returns:
        Dictionary with email components: to, subject, body, titles list
    """
    replacement_school = schools[replacement_code]

    # Get Chief Librarian info
    chief_name = replacement_school.get("chief_librarian_name", "")
    chief_email = replacement_school.get("chief_librarian_email", "")

    # Get first name for greeting
    first_name = chief_name.split()[0] if chief_name else "Colleague"

    # Build the titles section (each title with its Primo link for the replacement school)
    title_lines = []
    titles_info = []
    for result in titles_for_school:
        title = result.get("title", "Unknown title")
        nz_mms_id = result.get("bib_info", {}).get("nz_mms_id")
        primo_link = build_primo_link(nz_mms_id, replacement_school)

        if primo_link:
            title_lines.append(f"{title}\n{primo_link}")
        else:
            title_lines.append(title)

        titles_info.append({
            "barcode": result["barcode"],
            "title": title,
            "primo_link": primo_link
        })

    # Join titles with blank lines between them
    titles_section = "\n\n".join(title_lines)

    # Adjust wording based on number of titles
    if len(titles_for_school) == 1:
        titles_phrase = "this title"
    else:
        titles_phrase = "these titles"

    # Generate the email body
    body = f"""Hi {first_name},

Another CUNY library recently reported that they are no longer able to retain monographs they had previously committed to as part of our shared retention agreement. Rather than withdrawing the titles entirely from the consortium's retention pool, I'm reaching out to ask whether your library would be willing to take over the retention commitment for {titles_phrase}:

{titles_section}

If you're open to this, I'll update the relevant records accordingly. Please let me know if you'd be willing to assume this commitment or if you have any questions before deciding.

Thank you for considering this request, and for your continued support of our shared collections.

- Kristen"""

    return {
        "to": f"{chief_name} <{chief_email}>",
        "subject": "CUNY Shared Print retention commitment inquiry",
        "body": body,
        "replacement_school": replacement_school["name"],
        "replacement_code": replacement_code,
        "titles": titles_info
    }


def save_eml_file(email, output_dir):
    """
    Save a draft email as an .eml file that Outlook can open.

    Args:
        email: Dictionary with to, subject, body, replacement_school, titles
        output_dir: Directory to save the .eml file

    Returns:
        Path to the saved file
    """
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from datetime import datetime
    import os

    # Create the email message
    msg = MIMEMultipart('alternative')
    msg['To'] = email['to']
    msg['Subject'] = email['subject']
    msg['X-Unsent'] = '1'  # Tells Outlook this is a draft

    # Plain text version
    text_part = MIMEText(email['body'], 'plain', 'utf-8')

    # HTML version with clickable links
    html_body = email['body'].replace('\n', '<br>\n')
    # Make URLs clickable
    for t in email['titles']:
        if t['primo_link']:
            # Replace the plain URL with a hyperlink
            html_body = html_body.replace(
                t['primo_link'],
                f'<a href="{t["primo_link"]}">{t["title"]}</a>'
            )
            # Remove the title line since it's now in the link
            html_body = html_body.replace(f'{t["title"]}<br>', '')

    html_part = MIMEText(f'<html><body style="font-family: Arial, sans-serif;">{html_body}</body></html>', 'html', 'utf-8')

    msg.attach(text_part)
    msg.attach(html_part)

    # Create filename from school name
    safe_school_name = email['replacement_school'].replace(' ', '_').replace('/', '-')
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"retention_inquiry_{safe_school_name}_{timestamp}.eml"
    filepath = os.path.join(output_dir, filename)

    # Save the file
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(msg.as_string())

    return filepath


def print_draft_emails(results, schools, output_dir=None):
    """
    Generate draft emails, batching multiple titles per replacement school.
    Optionally save as .eml files.

    Args:
        results: List of result dictionaries from process_barcodes()
        schools: Dictionary of school data
        output_dir: If provided, save .eml files to this directory
    """
    # Filter to only items with replacements
    found = [r for r in results if r["status"] == "replacement_found"]

    if not found:
        print("\nNo draft emails to generate (no replacements found).")
        return []

    # Group by replacement school
    by_school = {}
    for result in found:
        school_code = result["replacement_school"]
        if school_code not in by_school:
            by_school[school_code] = []
        by_school[school_code].append(result)

    print("\n" + "=" * 60)
    print(f"DRAFT EMAILS ({len(by_school)} email(s) for {len(found)} title(s))")
    print("=" * 60)

    # Create output directory if needed
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    emails = []
    saved_files = []
    for i, (school_code, titles_for_school) in enumerate(by_school.items(), 1):
        email = generate_draft_email(titles_for_school, school_code, schools)
        emails.append(email)

        print(f"\n--- Email {i} of {len(by_school)} ---")
        print(f"To: {email['to']}")
        print(f"Subject: {email['subject']}")
        print(f"School: {email['replacement_school']}")
        print(f"Titles: {len(email['titles'])}")
        for t in email['titles']:
            print(f"  - {t['barcode']}: {t['title'][:50]}...")

        # Save as .eml file if output_dir provided
        if output_dir:
            filepath = save_eml_file(email, output_dir)
            saved_files.append(filepath)
            print(f"  Saved: {filepath}")
        else:
            print("-" * 40)
            print(email['body'])
            print("-" * 40)

    if saved_files:
        print(f"\n✓ Saved {len(saved_files)} .eml file(s) to: {output_dir}")
        print("  Double-click to open in Outlook as a draft message.")

    return emails


# =============================================================================
# PHASE 3: UPDATE LEAVING SCHOOL'S ALMA RECORDS
# =============================================================================

def update_leaving_school_item(mms_id, holding_id, item_pid, api_key, base_url):
    """
    Update the leaving school's item record:
    - Set committed_to_retain to false
    - Clear retention_reason

    Returns: (success, message)
    """
    url = f"{base_url}/almaws/v1/bibs/{mms_id}/holdings/{holding_id}/items/{item_pid}"
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    params = {"apikey": api_key}

    # GET current item record
    response = requests.get(url, params=params, headers=headers)
    if response.status_code != 200:
        return False, f"GET item failed (status {response.status_code}): {response.text}"

    item = response.json()

    # Modify retention fields
    item["item_data"]["committed_to_retain"] = {"value": "false"}
    item["item_data"]["retention_reason"] = {"value": ""}

    # PUT updated item record
    put_response = requests.put(url, params=params, headers=headers, json=item)
    if put_response.status_code != 200:
        return False, f"PUT item failed (status {put_response.status_code}): {put_response.text}"

    return True, "Item updated: committed_to_retain=No, retention_reason cleared"


def get_other_retained_items(mms_id, holding_id, current_item_pid, api_key, base_url):
    """
    Get all items under a holding that still have committed_to_retain=true,
    excluding the current item being processed.

    Returns: list of item PIDs with committed_to_retain=true (excluding current item)
    """
    url = f"{base_url}/almaws/v1/bibs/{mms_id}/holdings/{holding_id}/items"
    headers = {"Accept": "application/json"}
    params = {"apikey": api_key, "limit": 100}

    response = requests.get(url, params=params, headers=headers)
    if response.status_code != 200:
        return None  # Error - can't determine, skip holdings update

    data = response.json()
    items = data.get("item", [])

    retained = []
    for item in items:
        pid = item.get("item_data", {}).get("pid", "")
        if pid == current_item_pid:
            continue  # Skip current item
        committed = item.get("item_data", {}).get("committed_to_retain", {}).get("value", "")
        if committed == "true":
            retained.append(pid)

    return retained


def update_leaving_school_holdings(mms_id, holding_id, item_pid, api_key, base_url):
    """
    Update the leaving school's holdings record:
    - Remove the MARC 583 field, but only if no other retained items remain under this holding.

    Returns: (success, message)
    """
    # Check for other retained items under this holding
    other_retained = get_other_retained_items(mms_id, holding_id, item_pid, api_key, base_url)

    if other_retained is None:
        return False, "Could not check other items under this holding - skipping holdings update"

    if other_retained:
        return True, f"Holdings 583 field left in place ({len(other_retained)} other retained item(s) remain under this holding)"

    # No other retained items - safe to remove 583 field
    url = f"{base_url}/almaws/v1/bibs/{mms_id}/holdings/{holding_id}"
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    params = {"apikey": api_key}

    # GET current holdings record
    response = requests.get(url, params=params, headers=headers)
    if response.status_code != 200:
        return False, f"GET holdings failed (status {response.status_code}): {response.text}"

    holdings = response.json()

    # Remove 583 fields from MARC XML in the 'anies' field
    marc_xml = holdings.get("anies", [""])[0]
    if not marc_xml:
        return True, "No MARC XML found in holdings - nothing to update"

    # Remove all 583 datafield elements
    updated_xml = re.sub(r'<datafield tag="583"[^>]*>.*?</datafield>', '', marc_xml, flags=re.DOTALL)

    if updated_xml == marc_xml:
        return True, "No 583 field found in holdings - nothing to remove"

    holdings["anies"] = [updated_xml]

    # PUT updated holdings record
    put_response = requests.put(url, params=params, headers=headers, json=holdings)
    if put_response.status_code != 200:
        return False, f"PUT holdings failed (status {put_response.status_code}): {put_response.text}"

    return True, "Holdings updated: 583 field removed"


def update_leaving_school(result, schools, config):
    """
    Orchestrate Phase 3 updates for one item's leaving school.

    Returns: dict with update results
    """
    leaving_school_code = result["leaving_school"]
    school = schools[leaving_school_code]
    api_key = school.get("api_key", "")
    base_url = config["base_url"]

    item_data = result.get("bib_info", {}).get("item_data", {})
    mms_id = item_data.get("bib_data", {}).get("mms_id", "")
    holding_id = item_data.get("holding_data", {}).get("holding_id", "")
    item_pid = item_data.get("item_data", {}).get("pid", "")
    title = result.get("title", "Unknown")

    print(f"\n  Updating leaving school records for: {title[:60]}")
    print(f"  School: {school['name']}")

    if not all([mms_id, holding_id, item_pid, api_key]):
        print("  ERROR: Missing required IDs or API key - skipping")
        return {"status": "error", "message": "Missing IDs or API key"}

    # Update item record
    item_ok, item_msg = update_leaving_school_item(mms_id, holding_id, item_pid, api_key, base_url)
    if item_ok:
        print(f"  ✓ {item_msg}")
    else:
        print(f"  ✗ Item update failed: {item_msg}")
        return {"status": "error", "message": item_msg}

    # Update holdings record
    holdings_ok, holdings_msg = update_leaving_school_holdings(mms_id, holding_id, item_pid, api_key, base_url)
    if holdings_ok:
        print(f"  ✓ {holdings_msg}")
    else:
        print(f"  ✗ Holdings update failed: {holdings_msg}")
        return {"status": "partial", "message": f"Item updated but holdings failed: {holdings_msg}"}

    return {"status": "success", "message": "Item and holdings updated"}


def process_leaving_school_updates(results, schools, config):
    """
    Phase 3: Update leaving school records for all items where a replacement was found.
    Asks for confirmation before making any changes.
    """
    found = [r for r in results if r["status"] == "replacement_found"]

    if not found:
        print("\nNo items to update.")
        return

    print("\n" + "=" * 60)
    print(f"PHASE 3: UPDATE LEAVING SCHOOL RECORDS ({len(found)} item(s))")
    print("=" * 60)

    if config.get("sandbox"):
        print("  ⚠️  SANDBOX MODE - changes will be made to sandbox only")
    else:
        print("  ⚠️  PRODUCTION MODE - changes will be made to live Alma records")

    print("\nItems to update:")
    for r in found:
        leaving_name = schools.get(r["leaving_school"], {}).get("name", r["leaving_school"])
        print(f"  - {r['title'][:60]}")
        print(f"    Barcode: {r['barcode']} | School: {leaving_name}")

    confirm = input(f"\nUpdate leaving school records for {len(found)} item(s)? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Skipping Phase 3 updates.")
        return

    update_results = []
    for result in found:
        update = update_leaving_school(result, schools, config)
        update_results.append({"barcode": result["barcode"], "title": result["title"], **update})

    # Summary
    succeeded = [u for u in update_results if u["status"] == "success"]
    partial = [u for u in update_results if u["status"] == "partial"]
    failed = [u for u in update_results if u["status"] == "error"]

    print("\n--- Phase 3 Summary ---")
    print(f"  ✓ Fully updated: {len(succeeded)}")
    print(f"  ~ Partially updated: {len(partial)}")
    print(f"  ✗ Failed: {len(failed)}")

    if failed:
        print("\nFailed items:")
        for u in failed:
            print(f"  - {u['barcode']}: {u['message']}")

    return update_results


# =============================================================================
# PHASE 4: UPDATE TAKING SCHOOL'S ALMA RECORDS
# =============================================================================

# 583 field constants for CUNY Shared Print retention
RETENTION_583_START_DATE = "20241001"   # $c  Program start date (fixed)
RETENTION_583_END_DATE   = "20380930"   # $d  Program end date (fixed)
RETENTION_583_ORG_NAME   = "CUNY"       # $f  Consortium name
RETENTION_583_SOURCE     = "pda"        # $2  Source/scheme


def get_taking_school_iz_mms_id(nz_mms_id, api_key, base_url):
    """
    Look up the taking school's Institution Zone MMS ID using the NZ MMS ID.

    Calls: GET /almaws/v1/bibs?nz_mms_id={nz_mms_id}  with the taking school's IZ API key.
    Alma returns the corresponding IZ bib record for that school.

    Args:
        nz_mms_id: Network Zone MMS ID (from Phase 1 lookup)
        api_key:   Taking school's Institution Zone API key
        base_url:  Alma API base URL

    Returns:
        IZ MMS ID string, or None if not found
    """
    url = f"{base_url}/almaws/v1/bibs"
    headers = {"Accept": "application/json"}
    params = {"nz_mms_id": nz_mms_id, "apikey": api_key}

    try:
        response = requests.get(url, params=params, headers=headers)
        if response.status_code != 200:
            return None, f"GET bib by NZ MMS ID failed (status {response.status_code}): {response.text}"

        data = response.json()
        bibs = data.get("bib", [])
        if not bibs:
            return None, "No IZ bib record found for this NZ MMS ID at the taking school"

        iz_mms_id = bibs[0].get("mms_id")
        return iz_mms_id, None

    except requests.RequestException as e:
        return None, f"Connection error: {e}"


def get_taking_school_holding_and_item(iz_mms_id, api_key, base_url):
    """
    Find a holding and item record at the taking school for an IZ MMS ID.

    If the bib has exactly one item across all holdings, return it.
    If it has more than one item (e.g. multiple volumes or copies), return a
    "needs_review" error — we cannot safely determine which item to update
    without knowing the specific volume/copy the taking school is agreeing to retain.

    Args:
        iz_mms_id: Taking school's Institution Zone MMS ID
        api_key:   Taking school's IZ API key
        base_url:  Alma API base URL

    Returns:
        (holding_id, item_pid, error_message)
        holding_id and item_pid are None if not found or ambiguous.
        error_message begins with "NEEDS_REVIEW:" if manual intervention is needed.
    """
    # GET all holdings for this bib
    holdings_url = f"{base_url}/almaws/v1/bibs/{iz_mms_id}/holdings"
    headers = {"Accept": "application/json"}
    params = {"apikey": api_key}

    try:
        response = requests.get(holdings_url, params=params, headers=headers)
        if response.status_code != 200:
            return None, None, f"GET holdings failed (status {response.status_code}): {response.text}"

        data = response.json()
        holdings = data.get("holding", [])
        if not holdings:
            return None, None, "No holdings found at taking school"

        # Collect all items across all holdings
        all_items = []  # list of (holding_id, item_pid, enumeration)
        for holding in holdings:
            holding_id = holding.get("holding_id")
            if not holding_id:
                continue

            items_url = f"{base_url}/almaws/v1/bibs/{iz_mms_id}/holdings/{holding_id}/items"
            items_response = requests.get(items_url, params=params, headers=headers)
            if items_response.status_code != 200:
                continue

            items_data = items_response.json()
            for item in items_data.get("item", []):
                item_pid = item.get("item_data", {}).get("pid")
                enumeration = item.get("item_data", {}).get("enumeration_a", "")
                if item_pid:
                    all_items.append((holding_id, item_pid, enumeration))

        if not all_items:
            return None, None, "No items found in any holding at taking school"

        if len(all_items) > 1:
            # Multiple items — cannot safely pick one without enumeration matching
            descriptions = ", ".join(
                f"item {pid} ({enum})" if enum else f"item {pid}"
                for _, pid, enum in all_items
            )
            return None, None, f"NEEDS_REVIEW: {len(all_items)} items found ({descriptions}) — cannot determine which to update without enumeration matching"

        # Exactly one item — safe to proceed
        holding_id, item_pid, _ = all_items[0]
        return holding_id, item_pid, None

    except requests.RequestException as e:
        return None, None, f"Connection error: {e}"


def build_583_field(marc_org_code):
    """
    Build the MARC XML for a 583 retention action note field.

    Format:
        583 1_ $a Committed to retain $c 20241001 $d 20380930 $f CUNY $2 pda $5 <marc_org_code>

    The 583 field uses indicators "1" and " " (blank).

    Args:
        marc_org_code: The taking school's MARC organization code (e.g., "NyNyBC")

    Returns:
        XML string for the 583 datafield element
    """
    # Build subfields
    subfields = [
        f'<subfield code="a">Committed to retain</subfield>',
        f'<subfield code="c">{RETENTION_583_START_DATE}</subfield>',
        f'<subfield code="d">{RETENTION_583_END_DATE}</subfield>',
        f'<subfield code="f">{RETENTION_583_ORG_NAME}</subfield>',
        f'<subfield code="2">{RETENTION_583_SOURCE}</subfield>',
    ]
    # Only include $5 if a MARC org code is available
    if marc_org_code:
        subfields.append(f'<subfield code="5">{marc_org_code}</subfield>')
    subfields_xml = "\n    ".join(subfields)
    return f'<datafield tag="583" ind1="1" ind2=" ">\n    {subfields_xml}\n  </datafield>'


def update_taking_school_item(iz_mms_id, holding_id, item_pid, api_key, base_url):
    """
    Update the taking school's item record:
    - Set committed_to_retain to true
    - Set retention_reason to CUNYSharedPrint

    Returns: (success, message)
    """
    url = f"{base_url}/almaws/v1/bibs/{iz_mms_id}/holdings/{holding_id}/items/{item_pid}"
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    params = {"apikey": api_key}

    # GET current item record
    response = requests.get(url, params=params, headers=headers)
    if response.status_code != 200:
        return False, f"GET item failed (status {response.status_code}): {response.text}"

    item = response.json()

    # Modify retention fields
    item["item_data"]["committed_to_retain"] = {"value": "true"}
    item["item_data"]["retention_reason"] = {"value": "CUNYSharedPrint"}

    # PUT updated item record
    put_response = requests.put(url, params=params, headers=headers, json=item)
    if put_response.status_code != 200:
        return False, f"PUT item failed (status {put_response.status_code}): {put_response.text}"

    return True, "Item updated: committed_to_retain=Yes, retention_reason=CUNYSharedPrint"


def update_taking_school_holdings(iz_mms_id, holding_id, marc_org_code, api_key, base_url):
    """
    Update the taking school's holdings record by adding a MARC 583 field.

    Only adds the field if one doesn't already exist (avoids duplicates).

    Args:
        iz_mms_id:     Taking school's IZ MMS ID
        holding_id:    Holdings record ID
        marc_org_code: Taking school's MARC organization code (for $5)
        api_key:       Taking school's IZ API key
        base_url:      Alma API base URL

    Returns: (success, message)
    """
    url = f"{base_url}/almaws/v1/bibs/{iz_mms_id}/holdings/{holding_id}"
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    params = {"apikey": api_key}

    # GET current holdings record
    response = requests.get(url, params=params, headers=headers)
    if response.status_code != 200:
        return False, f"GET holdings failed (status {response.status_code}): {response.text}"

    holdings = response.json()
    marc_xml = holdings.get("anies", [""])[0]
    if not marc_xml:
        return False, "No MARC XML found in holdings record"

    # Check if a 583 field already exists
    if re.search(r'<datafield tag="583"', marc_xml):
        return True, "Holdings already has a 583 field - no change needed"

    # Build the new 583 field XML
    new_583 = build_583_field(marc_org_code)

    # Insert the 583 field before the closing </record> tag
    if "</record>" not in marc_xml:
        return False, "Could not find </record> tag in holdings MARC XML - cannot insert 583"

    updated_xml = marc_xml.replace("</record>", f"  {new_583}\n</record>")

    holdings["anies"] = [updated_xml]

    # PUT updated holdings record
    put_response = requests.put(url, params=params, headers=headers, json=holdings)
    if put_response.status_code != 200:
        return False, f"PUT holdings failed (status {put_response.status_code}): {put_response.text}"

    return True, "Holdings updated: 583 field added"


def update_taking_school(result, schools, config):
    """
    Orchestrate Phase 4 updates for one item's taking school.

    Steps:
    1. Look up the taking school's IZ MMS ID from the NZ MMS ID
    2. Find the holding and item at the taking school
    3. Update the item record (committed_to_retain, retention_reason)
    4. Update the holdings record (add 583 field)

    Returns: dict with update results, and stores IZ IDs back onto result
             so Phase 5 (WorldCat) can use them.
    """
    taking_school_code = result.get("replacement_school")
    if not taking_school_code:
        return {"status": "error", "message": "No replacement school recorded"}

    school = schools.get(taking_school_code)
    if not school:
        return {"status": "error", "message": f"School '{taking_school_code}' not in schools list"}

    api_key = school.get("api_key", "")
    marc_org_code = school.get("marc_org_code", "")
    base_url = config["base_url"]

    nz_mms_id = result.get("bib_info", {}).get("nz_mms_id")
    title = result.get("title", "Unknown")

    print(f"\n  Updating taking school records for: {title[:60]}")
    print(f"  School: {school['name']}")

    if not api_key or pd.isna(api_key):
        print("  ERROR: No API key for taking school - skipping")
        return {"status": "error", "message": "No API key for taking school"}

    try:
        marc_org_code_missing = not marc_org_code or pd.isna(marc_org_code) or str(marc_org_code).lower() == "nan"
    except (TypeError, ValueError):
        marc_org_code_missing = True
    if marc_org_code_missing:
        print("  WARNING: No MARC Org Code for taking school - 583 $5 subfield will be empty")
        marc_org_code = ""

    if not nz_mms_id:
        print("  ERROR: No NZ MMS ID available - skipping")
        return {"status": "error", "message": "No NZ MMS ID"}

    # Step 1: Get taking school's IZ MMS ID
    iz_mms_id, err = get_taking_school_iz_mms_id(nz_mms_id, api_key, base_url)
    if err:
        print(f"  ✗ Could not find IZ bib record: {err}")
        return {"status": "error", "message": err}
    print(f"  Found IZ MMS ID: {iz_mms_id}")

    # Step 2: Find holding and item
    holding_id, item_pid, err = get_taking_school_holding_and_item(iz_mms_id, api_key, base_url)
    if err:
        if err.startswith("NEEDS_REVIEW:"):
            print(f"  ⚠️  Manual review required: {err[len('NEEDS_REVIEW:'):].strip()}")
            return {"status": "needs_review", "message": err[len("NEEDS_REVIEW:"):].strip()}
        print(f"  ✗ Could not find holding/item: {err}")
        return {"status": "error", "message": err}
    print(f"  Found holding: {holding_id}, item: {item_pid}")

    # Store IDs on result for Phase 5 (WorldCat)
    result.setdefault("taking_school_ids", {})[taking_school_code] = {
        "iz_mms_id": iz_mms_id,
        "holding_id": holding_id,
        "item_pid": item_pid
    }

    # Step 3: Update item record
    item_ok, item_msg = update_taking_school_item(iz_mms_id, holding_id, item_pid, api_key, base_url)
    if item_ok:
        print(f"  ✓ {item_msg}")
    else:
        print(f"  ✗ Item update failed: {item_msg}")
        return {"status": "error", "message": item_msg}

    # Step 4: Update holdings record
    holdings_ok, holdings_msg = update_taking_school_holdings(
        iz_mms_id, holding_id, marc_org_code, api_key, base_url
    )
    if holdings_ok:
        print(f"  ✓ {holdings_msg}")
    else:
        print(f"  ✗ Holdings update failed: {holdings_msg}")
        return {"status": "partial", "message": f"Item updated but holdings failed: {holdings_msg}"}

    return {"status": "success", "message": "Item and holdings updated"}


def process_taking_school_updates(results, schools, config):
    """
    Phase 4: Update taking school records for all items where a replacement was found
    AND Phase 3 succeeded (or was run).

    Asks for user confirmation before making any changes.
    """
    # Only process items that have a replacement school
    found = [r for r in results if r["status"] == "replacement_found"]

    if not found:
        print("\nNo items to update for taking school.")
        return []

    print("\n" + "=" * 60)
    print(f"PHASE 4: UPDATE TAKING SCHOOL RECORDS ({len(found)} item(s))")
    print("=" * 60)

    if config.get("sandbox"):
        print("  ⚠️  SANDBOX MODE - changes will be made to sandbox only")
    else:
        print("  ⚠️  PRODUCTION MODE - changes will be made to live Alma records")

    print("\nItems to update:")
    for r in found:
        taking_name = schools.get(r["replacement_school"], {}).get("name", r["replacement_school"])
        print(f"  - {r['title'][:60]}")
        print(f"    Barcode: {r['barcode']} → Taking school: {taking_name}")

    confirm = input(f"\nUpdate taking school records for {len(found)} item(s)? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Skipping Phase 4 updates.")
        return []

    update_results = []
    for result in found:
        update = update_taking_school(result, schools, config)
        update_results.append({"barcode": result["barcode"], "title": result["title"], **update})

    # Summary
    succeeded = [u for u in update_results if u["status"] == "success"]
    partial = [u for u in update_results if u["status"] == "partial"]
    needs_review = [u for u in update_results if u["status"] == "needs_review"]
    failed = [u for u in update_results if u["status"] == "error"]

    print("\n--- Phase 4 Summary ---")
    print(f"  ✓ Fully updated: {len(succeeded)}")
    print(f"  ~ Partially updated: {len(partial)}")
    print(f"  ⚠️  Needs manual review: {len(needs_review)}")
    print(f"  ✗ Failed: {len(failed)}")

    if needs_review:
        print("\nItems requiring manual review (multiple volumes/copies found):")
        for u in needs_review:
            print(f"  - {u['barcode']}: {u['title'][:60]}")
            print(f"    {u['message']}")

    if failed:
        print("\nFailed items:")
        for u in failed:
            print(f"  - {u['barcode']}: {u['message']}")

    return update_results


# =============================================================================
# MAIN WORKFLOW
# =============================================================================

def read_barcodes(file_path):
    """
    Read barcodes and school codes from Excel file.

    Expected columns:
        - Barcode: The item barcode
        - School Code: The Alma Institution Code (e.g., 01CUNY_QC)

    Returns a list of tuples: [(barcode, school_code), ...]
    """
    print(f"Reading barcodes from: {file_path}")

    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    if len(df.columns) < 2:
        print("ERROR: Excel file must have two columns: Barcode and School Code")
        print("       Example: Barcode | School Code")
        print("                31699001195116 | 01CUNY_QC")
        sys.exit(1)

    # Use first column as barcode, second column as school code
    barcode_col = df.columns[0]
    school_col = df.columns[1]

    print(f"  Barcode column: '{barcode_col}'")
    print(f"  School code column: '{school_col}'")

    # Build list of (barcode, school_code) tuples
    items = []
    for _, row in df.iterrows():
        barcode = str(row[barcode_col]).strip()
        school_code = str(row[school_col]).strip()

        if barcode and barcode.lower() != 'nan' and school_code and school_code.lower() != 'nan':
            items.append((barcode, school_code))

    print(f"Found {len(items)} items to process")
    return items


def process_barcodes(items, schools, config):
    """
    Process each barcode: look up holdings and select replacement school.

    Args:
        items: List of (barcode, school_code) tuples
        schools: Dictionary of school data
        config: Configuration dictionary

    Returns a list of results for each barcode.
    """
    results = []
    total = len(items)

    for i, (barcode, leaving_school) in enumerate(items, 1):
        print(f"\n[{i}/{total}] Processing barcode: {barcode}")

        # Validate leaving school
        if leaving_school not in schools:
            print(f"  ERROR: Unknown school code '{leaving_school}'")
            results.append({
                "barcode": barcode,
                "status": "error",
                "title": None,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "error": f"Unknown school code: {leaving_school}"
            })
            continue

        leaving_school_name = schools[leaving_school]["name"]
        print(f"  Leaving school: {leaving_school_name}")

        # Find which institutions hold this item
        institutions, bib_info = find_holding_institutions(
            barcode, leaving_school, schools, config
        )

        if institutions is None:
            print(f"  NOT FOUND in Alma")
            results.append({
                "barcode": barcode,
                "status": "not_found",
                "title": None,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": []
            })
            continue

        # Check if item is ineligible due to status
        if institutions == "ineligible":
            title = bib_info.get("title", "Unknown") if bib_info else "Unknown"
            status_desc = bib_info.get("status", "Unknown status")
            results.append({
                "barcode": barcode,
                "status": "ineligible",
                "title": title,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "item_status": status_desc,
                "bib_info": bib_info
            })
            continue

        title = bib_info.get("title", "Unknown") if bib_info else "Unknown"
        print(f"  Title: {title[:60]}...")
        print(f"  Held by: {len(institutions)} institution(s)")

        # Select replacement school
        replacement, eligible_list = select_replacement_school(
            institutions, leaving_school, schools
        )

        if replacement is None:
            print(f"  ⚠️  NO ELIGIBLE REPLACEMENT - Flag for withdrawal review")
            results.append({
                "barcode": barcode,
                "status": "no_replacement",
                "title": title,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "bib_info": bib_info
            })
        else:
            school_name = schools[replacement]["name"]
            print(f"  ✓ Recommended replacement: {school_name}")
            results.append({
                "barcode": barcode,
                "status": "replacement_found",
                "title": title,
                "leaving_school": leaving_school,
                "replacement_school": replacement,
                "eligible_schools": eligible_list,
                "bib_info": bib_info
            })

    return results


def print_summary(results, schools):
    """Print a summary of the results."""
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)

    found = [r for r in results if r["status"] == "replacement_found"]
    no_replacement = [r for r in results if r["status"] == "no_replacement"]
    not_found = [r for r in results if r["status"] == "not_found"]
    ineligible = [r for r in results if r["status"] == "ineligible"]

    print(f"\nReplacement found:       {len(found)}")
    print(f"No replacement (review): {len(no_replacement)}")
    print(f"Ineligible (wrong status): {len(ineligible)}")
    print(f"Not found in Alma:       {len(not_found)}")

    if found:
        print("\n--- Items with Replacements ---")
        for r in found:
            leaving_name = schools.get(r["leaving_school"], {}).get("name", "Unknown")
            replacement_name = schools[r["replacement_school"]]["name"]
            print(f"  {r['barcode']}")
            print(f"    Title: {r['title'][:50]}...")
            print(f"    From: {leaving_name} → To: {replacement_name}")

    if no_replacement:
        print("\n--- Items Flagged for Withdrawal Review ---")
        for r in no_replacement:
            leaving_name = schools.get(r["leaving_school"], {}).get("name", "Unknown")
            print(f"  {r['barcode']}")
            print(f"    Title: {r['title'][:50]}...")
            print(f"    From: {leaving_name} (no other schools hold this)")

    if ineligible:
        print("\n--- Ineligible Items (Item Not In Place) ---")
        for r in ineligible:
            leaving_name = schools.get(r["leaving_school"], {}).get("name", "Unknown")
            status = r.get("item_status", "Unknown status")
            print(f"  {r['barcode']}")
            print(f"    Title: {r['title'][:50]}...")
            print(f"    Status: {status}")

    if not_found:
        print("\n--- Barcodes Not Found ---")
        for r in not_found:
            print(f"  {r['barcode']}")

    print("=" * 60)

    return found, no_replacement, not_found, ineligible


def main():
    """Main function."""
    # Check arguments
    # Parse arguments - allow --sandbox flag anywhere
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    sandbox = "--sandbox" in sys.argv

    if len(args) < 1:
        print("Usage: python retention_transfer.py <barcodes.xlsx> [output_dir] [--sandbox]")
        print("\nThe Excel file should have two columns:")
        print("  - Barcode: The item barcode")
        print("  - School Code: The Alma Institution Code (e.g., 01CUNY_QC)")
        print("\nOptions:")
        print("  output_dir  Directory to save .eml files (default: output/emails)")
        print("  --sandbox   Use sandbox API keys instead of production")
        print("\nExamples:")
        print("  python retention_transfer.py data/barcodes.xlsx")
        print("  python retention_transfer.py data/barcodes.xlsx output/emails --sandbox")
        sys.exit(1)

    barcode_file = args[0]
    output_dir = args[1] if len(args) > 1 else "output/emails"

    print("=" * 60)
    print("CUNY Shared Print - Retention Transfer")
    print("=" * 60)

    # Load configuration
    config = get_config(sandbox=sandbox)
    if sandbox:
        print("⚠️  SANDBOX MODE")
    print(f"API Base URL: {config['base_url']}")

    # Load schools data
    schools = load_schools(config["schools_file"])
    print(f"Loaded {len(schools)} schools")

    # Read barcodes and school codes
    items = read_barcodes(barcode_file)

    if not items:
        print("No items found. Exiting.")
        sys.exit(1)

    # Phase 1: Look up items and select replacement schools
    print("\nLooking up items and selecting replacement schools...")
    results = process_barcodes(items, schools, config)

    # Print summary
    found, no_replacement, not_found, ineligible = print_summary(results, schools)

    # Phase 2: Generate draft emails and save as .eml files
    emails = print_draft_emails(results, schools, output_dir)

    # Phase 3: Update leaving school's Alma records
    process_leaving_school_updates(results, schools, config)

    # Phase 4: Update taking school's Alma records
    process_taking_school_updates(results, schools, config)

    # TODO Phase 5: Update WorldCat holdings


if __name__ == "__main__":
    main()
