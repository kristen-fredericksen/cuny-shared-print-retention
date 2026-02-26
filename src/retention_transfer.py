"""
CUNY Shared Print Monograph Trust - Retention Transfer Script

This script transfers retention commitments from one school to another
when a school can no longer retain a book. It runs in two separate steps
to enforce the email-and-wait workflow:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 1 — LOOKUP  (run once, right away)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    python retention_transfer.py data/barcodes.xlsx [--sandbox]

  - Reads barcodes from Excel, looks up each item in Alma
  - Selects a replacement school for each item
  - Generates draft .eml email files (one per replacement school)
  - Saves a pending-transfers JSON file and then STOPS
  - No Alma or WorldCat records are changed

  The Excel file must have two columns:
    Barcode      — the item barcode (e.g. 39016013760757)
    School Code  — the leaving school's Alma code (e.g. 01CUNY_QC)

  After this step: send the draft emails and wait for replies.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 2 — UPDATE  (run after you receive replies)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    python retention_transfer.py output/pending_YYYYMMDD_HHMMSS.json --update [--sandbox]

  For each pending item, prompts: "Did [school] agree? (yes/no/skip)"
    yes   — re-queries Alma to verify IDs, then updates Alma and generates
            WorldCat CSV files (Phases 3, 4, 5)
    no    — marks that school as declined, moves to the next eligible school,
            generates a new draft email, saves updated pending file
    skip  — leaves the item pending, tries again next run

  If all eligible schools decline, the item is flagged for withdrawal review.
  The pending file is updated after each run so you can re-run as needed.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Options
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  --sandbox   Use sandbox API keys and sandbox schools file
"""

import os
import sys
import re
import json
import requests
import pandas as pd
from datetime import datetime
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

def _clean_oclc_collection_id(raw):
    """
    Clean an OCLC Collection ID value read from a CSV/Excel cell.

    pandas reads a blank numeric column as NaN, and a present integer like
    1055226 as the float 1055226.0.  We want the clean string "1055226".

    Returns "" for blank/NaN values, otherwise the integer string.
    """
    if raw is None:
        return ""
    if isinstance(raw, float):
        if pd.isna(raw):
            return ""
        return str(int(raw))
    cleaned = str(raw).strip()
    if cleaned.lower() in ("", "nan"):
        return ""
    # Handle a string like "1055226.0" produced by earlier str() conversion
    try:
        return str(int(float(cleaned)))
    except ValueError:
        return cleaned


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
            "api_key": row.get("Alma API Key", ""),
            "oclc_collection_id": _clean_oclc_collection_id(row.get("OCLC Collection ID", ""))
        }

    return schools


def get_grad_center_code(schools):
    """
    Return the Grad Center's institution code.

    We match on the known stable institution code rather than the name,
    which could change or match other institutions unintentionally.
    """
    GRAD_CENTER_CODE = "01CUNY_GC"
    if GRAD_CENTER_CODE in schools:
        return GRAD_CENTER_CODE
    # Fallback: name-based search for sandbox or unusual configs
    for code, school in schools.items():
        if "grad" in school["name"].lower() and "center" in school["name"].lower():
            return code
    return None


# =============================================================================
# ALMA API FUNCTIONS
# =============================================================================

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

    # Step 4: Verify the leaving school actually holds this item.
    # If it's not in the NZ holdings list, the barcode/school pairing
    # in the spreadsheet is wrong — flag it rather than silently proceeding.
    if institutions is not None and leaving_school not in institutions:
        print(f"  ⚠️  MISMATCH: {leaving_school} is listed as the leaving school "
              f"but does not appear in the NZ holdings for this item.")
        print(f"  Actual holders: {institutions}")
        bib_info["actual_holders"] = institutions
        return "mismatch", bib_info

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
    # Filter to only Shared Print participants who hold the item,
    # tracking why each non-eligible holder was excluded.
    eligible = []
    only_leaving = True       # all holders are the leaving school
    non_participant_names = [] # holders that aren't in Shared Print

    for inst_code in holding_institutions:
        if inst_code == leaving_school:
            continue

        only_leaving = False  # at least one other school holds it

        if inst_code not in schools:
            continue  # School not in our schools file

        school = schools[inst_code]
        if not school["shared_print"]:
            non_participant_names.append(school["name"])
            continue

        eligible.append(school)

    if not eligible:
        # Build a human-readable reason for the caller / UI
        if only_leaving and not non_participant_names:
            reason = "No other CUNY school holds this item"
        elif non_participant_names:
            names = ", ".join(non_participant_names)
            reason = (f"Held by {names}, but they are not Shared Print "
                      f"participant(s)")
        else:
            reason = "No eligible Shared Print participants hold this item"
        return None, [], reason

    # Check if Grad Center is in the eligible list
    grad_center_code = get_grad_center_code(schools)
    for school in eligible:
        if school["code"] == grad_center_code:
            # Put Grad Center first, then sort rest by size
            others = [s for s in eligible if s["code"] != grad_center_code]
            others.sort(key=lambda s: s["size"])  # 1 = largest, so ascending
            priority_list = [school] + others
            return school["code"], [s["code"] for s in priority_list], ""

    # No Grad Center - sort by size (1 = largest)
    eligible.sort(key=lambda s: s["size"])

    return eligible[0]["code"], [s["code"] for s in eligible], ""


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

    Before making any change, verifies that committed_to_retain is currently
    true.  If it is already false, warns and skips (avoids silently "removing"
    a commitment that was never set).

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

    # Check that this item actually has a retention commitment to remove
    current_committed = item.get("item_data", {}).get("committed_to_retain", {}).get("value", "")
    if current_committed != "true":
        return False, (
            f"Item does not appear to have an active retention commitment "
            f"(committed_to_retain='{current_committed}'). "
            f"Skipping to avoid clearing a commitment that was never set."
        )

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

    # Warn if any taking school has no API key. If Phase 3 removes the leaving
    # school's commitment but Phase 4 cannot run for the taking school, the
    # item will have NO active retention commitment — an orphan.
    taking_no_key = []
    for r in found:
        taking_code = r.get("replacement_school")
        if taking_code:
            taking_school = schools.get(taking_code, {})
            api_key = taking_school.get("api_key", "")
            if not api_key or (isinstance(api_key, float) and pd.isna(api_key)):
                taking_no_key.append((r["barcode"], taking_school.get("name", taking_code)))
    if taking_no_key:
        print(f"\n⚠️  WARNING: {len(taking_no_key)} taking school(s) have no API key configured.")
        print("   If you proceed with Phase 3, the leaving school's commitment will be")
        print("   removed but Phase 4 will NOT be able to update the taking school.")
        print("   This would leave the item with NO active retention commitment.")
        for barcode, name in taking_no_key:
            print(f"   - Barcode {barcode}: taking school {name} has no API key")
        print("   Consider adding API keys for these schools before proceeding.")

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

    # Check if a CUNY Shared Print 583 field already exists.
    # We look for a 583 that contains both "Committed to retain" and the CUNY
    # org name, so we don't mistake an unrelated action note for our commitment.
    existing_583s = re.findall(
        r'<datafield tag="583"[^>]*>(.*?)</datafield>', marc_xml, flags=re.DOTALL
    )
    for field_xml in existing_583s:
        has_committed = "Committed to retain" in field_xml or "committed to retain" in field_xml
        has_cuny = RETENTION_583_ORG_NAME in field_xml
        if has_committed and has_cuny:
            return True, "Holdings already has a CUNY Shared Print 583 field - no change needed"

    # A 583 exists but it is NOT a CUNY Shared Print commitment — warn and proceed
    if existing_583s:
        print(f"  ⚠️  Holdings has {len(existing_583s)} existing 583 field(s) that are NOT CUNY Shared Print — adding ours anyway")

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
# PHASE 5: GENERATE WORLDCAT CSV FILES
# =============================================================================

# OCLC Full Format Shared Print File Template column headers (14 columns)
# These match the OCLC WorldShare Shared Print file specification.
WORLDCAT_CSV_HEADERS = [
    "OCLC Number",
    "LSN",
    "Library Symbol",
    "Collection ID",
    "ActionNote_583$a",
    "PublicNote_583$z",
    "ActionDate_583$c",
    "ExpirationDate_583$d",
    "Extent_of_commitment_583$n",
    "Method_of_determination_583$l",
    "Organization_583$f",
    "URI_583$u",
    "Copy_number_876-878$t",
    "Enumeration/chronology_876-878$3"
]


def get_oclc_number_from_item(item_data):
    """
    Extract the OCLC number from item data returned by Alma barcode lookup.

    Alma stores OCLC numbers in bib_data as 'oclc_number' or in the
    network_number list. The OCLC number may appear with a leading prefix
    like '(OCoLC)' which should be stripped.

    Returns:
        OCLC number string (digits only), or None if not found.
    """
    bib_data = item_data.get("bib_data", {})

    # Check direct oclc_number field first
    oclc_number = bib_data.get("oclc_number", "")
    if oclc_number:
        # Strip any prefix like "(OCoLC)"
        cleaned = re.sub(r'^\(OCoLC\)', '', str(oclc_number)).strip()
        if cleaned:
            return cleaned

    # Check network_number list for an OCLC entry
    network_numbers = bib_data.get("network_number", [])
    for nn in network_numbers:
        if "(OCoLC)" in nn:
            cleaned = re.sub(r'^\(OCoLC\)', '', nn).strip()
            if cleaned:
                return cleaned

    return None


def build_worldcat_row_taking(result, taking_school):
    """
    Build one CSV row for the TAKING school's WorldCat Shared Print submission.

    The taking school is ADDING a retention commitment.

    Args:
        result: Result dictionary from process_barcodes / Phase 4
        taking_school: School dictionary for the taking school

    Returns:
        List of values matching WORLDCAT_CSV_HEADERS order, or None if missing
        required data (OCLC number).
    """
    from datetime import datetime

    item_data = result.get("bib_info", {}).get("item_data", {})
    oclc_number = get_oclc_number_from_item(item_data)
    barcode = result.get("barcode", "")
    oclc_symbol = taking_school.get("oclc_symbol", "")
    collection_id = taking_school.get("oclc_collection_id", "")
    action_date = datetime.now().strftime("%Y%m%d")

    # OCLC number is required — without it the row cannot be processed
    if not oclc_number:
        return None

    return [
        oclc_number,                        # OCLC Number
        barcode,                            # LSN (Local System Number / barcode)
        oclc_symbol,                        # Library Symbol
        collection_id,                      # Collection ID
        "committed to retain",              # ActionNote_583$a
        "",                                 # PublicNote_583$z (not used)
        RETENTION_583_START_DATE,           # ActionDate_583$c (program start date)
        RETENTION_583_END_DATE,             # ExpirationDate_583$d (program end date)
        "",                                 # Extent_of_commitment_583$n (not used)
        "",                                 # Method_of_determination_583$l (not used)
        RETENTION_583_ORG_NAME,             # Organization_583$f
        "",                                 # URI_583$u (not used)
        "",                                 # Copy_number_876-878$t (not used)
        ""                                  # Enumeration/chronology_876-878$3 (not used)
    ]


def generate_worldcat_taking_csv(results, schools, output_dir):
    """
    Generate OCLC WorldCat CSV files for the TAKING school — adding retention commitments.

    Produces one CSV file per taking school (grouped by school).
    File naming: <collectionID>.<OCLCsymbol>.sharedprint_retention_transfer_<YYYYMMDD>.csv

    Files are saved to output_dir/worldcat/taking/.

    Args:
        results:    List of result dicts from process_barcodes (must have status=replacement_found)
        schools:    Dictionary of school data
        output_dir: Base output directory

    Returns:
        List of saved file paths.
    """
    import csv
    from datetime import datetime

    found = [r for r in results if r["status"] == "replacement_found"]
    if not found:
        return []

    # Group by taking school
    by_school = {}
    for result in found:
        taking_code = result.get("replacement_school")
        if not taking_code:
            continue
        if taking_code not in by_school:
            by_school[taking_code] = []
        by_school[taking_code].append(result)

    taking_dir = os.path.join(output_dir, "worldcat", "taking")
    os.makedirs(taking_dir, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    saved_files = []
    skipped_rows = []

    for school_code, school_results in by_school.items():
        school = schools.get(school_code, {})
        oclc_symbol = school.get("oclc_symbol", "")
        collection_id = school.get("oclc_collection_id", "")

        if not collection_id:
            print(f"  ⚠️  Skipping WorldCat CSV for {school.get('name', school_code)}: no OCLC Collection ID configured")
            skipped_rows.append({
                "school": school.get("name", school_code),
                "reason": "No OCLC Collection ID in schools CSV"
            })
            continue

        # Build filename
        filename = f"{collection_id}.{oclc_symbol}.sharedprint_retention_transfer_{today}.csv"
        filepath = os.path.join(taking_dir, filename)

        rows = []
        for result in school_results:
            row = build_worldcat_row_taking(result, school)
            if row is None:
                barcode = result.get("barcode", "?")
                title = result.get("title", "Unknown")
                print(f"  ⚠️  Skipping {barcode} ({title[:40]}): no OCLC number found")
                skipped_rows.append({
                    "school": school.get("name", school_code),
                    "barcode": barcode,
                    "reason": "No OCLC number found in Alma record"
                })
                continue
            rows.append(row)

        if not rows:
            print(f"  ⚠️  No valid rows for {school.get('name', school_code)} — CSV not saved")
            continue

        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(WORLDCAT_CSV_HEADERS)
            writer.writerows(rows)

        print(f"  ✓ Saved: {filepath} ({len(rows)} row(s))")
        saved_files.append(filepath)

    return saved_files, skipped_rows


def generate_worldcat_leaving_instructions(results, schools, output_dir):
    """
    For the LEAVING school, generate a plain-text instructions file listing
    the WorldCat LHRs that need to have their retention commitments removed.

    Removing a commitment can be done in two ways (per OCLC documentation):
    - Small batches: Use the LHR editor in WorldShare Record Manager to delete
      or edit retention commitments manually.
    - Large batches: Work with OCLC to set up a data sync collection to delete LHRs.

    This function creates a report file to assist with manual or batch removal.

    Args:
        results:    List of result dicts (status=replacement_found)
        schools:    Dictionary of school data
        output_dir: Base output directory

    Returns:
        Path to saved instructions file, or None if no items.
    """
    from datetime import datetime

    found = [r for r in results if r["status"] == "replacement_found"]
    if not found:
        return None

    leaving_dir = os.path.join(output_dir, "worldcat", "leaving")
    os.makedirs(leaving_dir, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    filename = f"worldcat_leaving_school_removal_instructions_{today}.txt"
    filepath = os.path.join(leaving_dir, filename)

    lines = []
    lines.append("WORLDCAT SHARED PRINT RETENTION REMOVAL INSTRUCTIONS")
    lines.append("=" * 60)
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append("The following items need their WorldCat Shared Print retention")
    lines.append("commitments REMOVED from the leaving school's WorldCat LHR.")
    lines.append("")
    lines.append("HOW TO REMOVE RETENTION COMMITMENTS IN WORLDCAT:")
    lines.append("")
    lines.append("Option A - Small batches (a few records):")
    lines.append("  Use the LHR editor in WorldShare Record Manager.")
    lines.append("  Search for each OCLC number, open the LHR, and delete or")
    lines.append("  edit the 583 retention commitment field.")
    lines.append("")
    lines.append("Option B - Large batches (many records):")
    lines.append("  Work with your OCLC database specialist or implementation")
    lines.append("  manager to create a data sync collection to delete LHRs.")
    lines.append("  When deleting, choose whether to keep or delete the WorldCat")
    lines.append("  holding on the bibliographic record.")
    lines.append("")
    lines.append("-" * 60)
    lines.append("ITEMS TO PROCESS:")
    lines.append("-" * 60)
    lines.append("")

    # Group by leaving school
    by_school = {}
    for result in found:
        leaving_code = result.get("leaving_school")
        if leaving_code not in by_school:
            by_school[leaving_code] = []
        by_school[leaving_code].append(result)

    for school_code, school_results in by_school.items():
        school = schools.get(school_code, {})
        school_name = school.get("name", school_code)
        oclc_symbol = school.get("oclc_symbol", "")
        lines.append(f"School: {school_name} (OCLC Symbol: {oclc_symbol})")
        lines.append("")

        for result in school_results:
            item_data = result.get("bib_info", {}).get("item_data", {})
            oclc_number = get_oclc_number_from_item(item_data) or "UNKNOWN - check Alma record"
            barcode = result.get("barcode", "")
            title = result.get("title", "Unknown")
            taking_code = result.get("replacement_school", "")
            taking_name = schools.get(taking_code, {}).get("name", taking_code)

            lines.append(f"  Title:       {title}")
            lines.append(f"  Barcode:     {barcode}")
            lines.append(f"  OCLC Number: {oclc_number}")
            lines.append(f"  Transferred to: {taking_name}")
            lines.append("")

        lines.append("")

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return filepath


def process_worldcat_updates(results, schools, config, output_dir):
    """
    Phase 5: Generate WorldCat CSV files and instructions for record updates.

    For the TAKING school: generates an OCLC Full Format Shared Print CSV
    to upload via WorldShare Metadata > My Files > Uploads (Data sync LHR).

    For the LEAVING school: generates an instructions text file explaining
    how to remove the retention commitments in WorldCat (via Record Manager
    for small batches, or data sync collection for large batches).

    Asks for confirmation before generating files.
    """
    found = [r for r in results if r["status"] == "replacement_found"]

    if not found:
        print("\nNo items to process for WorldCat updates.")
        return

    print("\n" + "=" * 60)
    print(f"PHASE 5: GENERATE WORLDCAT UPDATE FILES ({len(found)} item(s))")
    print("=" * 60)

    print("\nThis will generate:")
    print("  • CSV file(s) for TAKING school(s) — upload to OCLC WorldShare")
    print("  • Instructions file for LEAVING school(s) — remove old commitments")

    confirm = input(f"\nGenerate WorldCat update files for {len(found)} item(s)? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Skipping Phase 5.")
        return

    # Derive the base output directory from output_dir.
    # output_dir is typically "output/emails"; we want "output" as the base.
    # We use os.path.dirname to go up one level, but only if output_dir ends
    # with a recognised subdirectory name ("emails"). Otherwise use output_dir
    # itself so that a custom path like "output/run1" still works sensibly.
    output_dir_norm = os.path.normpath(output_dir)
    if os.path.basename(output_dir_norm) == "emails":
        worldcat_output = os.path.dirname(output_dir_norm)
    else:
        worldcat_output = output_dir_norm
    # Ensure we always have an actual path, not an empty string
    if not worldcat_output:
        worldcat_output = "output"

    # Generate taking school CSVs
    print("\n--- Taking school CSV files ---")
    saved_files, skipped = generate_worldcat_taking_csv(results, schools, worldcat_output)

    # Generate leaving school instructions
    print("\n--- Leaving school removal instructions ---")
    instructions_path = generate_worldcat_leaving_instructions(results, schools, worldcat_output)
    if instructions_path:
        print(f"  ✓ Saved: {instructions_path}")

    # Final summary
    print("\n--- Phase 5 Summary ---")
    if saved_files:
        print(f"  ✓ Taking school CSV files saved: {len(saved_files)}")
        for f in saved_files:
            print(f"    {f}")
        print()
        print("  Next step for TAKING school files:")
        print("  1. Log in to OCLC WorldShare Metadata")
        print("  2. Go to My Files > Uploads")
        print("  3. Select file type: 'Data sync LHR'")
        print("  4. Upload each CSV file")
    else:
        print("  No taking school CSV files were saved.")
        if skipped:
            print("  Skipped schools (missing OCLC Collection ID):")
            for s in skipped:
                print(f"    - {s['school']}: {s['reason']}")

    if instructions_path:
        print(f"\n  ✓ Leaving school instructions: {instructions_path}")
        print("  Next step for LEAVING school files:")
        print("  • Small batches: use WorldShare Record Manager LHR editor")
        print("  • Large batches: contact OCLC to set up a delete-LHR data sync collection")


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
        raw_barcode = row[barcode_col]
        school_code = str(row[school_col]).strip()

        # Excel sometimes converts long numeric barcodes to floats (e.g. 3.9E+13).
        # Convert float → int → string to recover the original digits.
        if isinstance(raw_barcode, float) and not pd.isna(raw_barcode):
            barcode = str(int(raw_barcode))
        else:
            barcode = str(raw_barcode).strip()

        if barcode and barcode.lower() != 'nan' and school_code and school_code.lower() != 'nan':
            items.append((barcode, school_code))

    # Warn about duplicate barcodes — processing the same barcode twice could
    # double-update records or generate duplicate emails.
    seen = {}
    for barcode, school_code in items:
        seen.setdefault(barcode, []).append(school_code)
    duplicates = {b: schools_list for b, schools_list in seen.items() if len(schools_list) > 1}
    if duplicates:
        print(f"\n⚠️  WARNING: {len(duplicates)} duplicate barcode(s) found in input file:")
        for barcode, schools_list in duplicates.items():
            print(f"  {barcode} appears {len(schools_list)} time(s)")
        print("  Each duplicate will be processed separately. Consider removing duplicates before proceeding.")

    print(f"Found {len(items)} items to process")
    return items


def process_barcodes(items, schools, config, progress_callback=None):
    """
    Process each barcode: look up holdings and select replacement school.

    Args:
        items: List of (barcode, school_code) tuples
        schools: Dictionary of school data
        config: Configuration dictionary
        progress_callback: Optional function called after each item with
                           (current_index, total, barcode, status_string).
                           Used by the Streamlit app to update a progress bar.

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
            if progress_callback:
                progress_callback(i, total, barcode, "error")
            continue

        leaving_school_name = schools[leaving_school]["name"]
        print(f"  Leaving school: {leaving_school_name}")

        # Warn if the leaving school is not a Shared Print participant.
        # This doesn't block processing, but it's probably a data entry error.
        if not schools[leaving_school].get("shared_print", False):
            print(f"  ⚠️  WARNING: {leaving_school_name} is not listed as a Shared Print participant."
                  f" Proceeding, but verify this barcode is correct.")

        # Find which institutions hold this item
        institutions, bib_info = find_holding_institutions(
            barcode, leaving_school, schools, config
        )

        # Barcode/school mismatch — the item exists but the leaving school
        # listed in the spreadsheet doesn't actually hold it
        if institutions == "mismatch":
            actual_holders = bib_info.get("actual_holders", []) if bib_info else []
            holder_names = ", ".join(
                schools.get(c, {}).get("name", c) for c in actual_holders
            ) or "unknown"
            title = bib_info.get("title", "Unknown") if bib_info else "Unknown"
            results.append({
                "barcode": barcode,
                "status": "error",
                "title": title,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "error": (
                    f"{leaving_school_name} is listed as the leaving school "
                    f"but does not hold this item in the Network Zone. "
                    f"Check that the school code in the spreadsheet is correct."
                ),
                "bib_info": bib_info,
            })
            if progress_callback:
                progress_callback(i, total, barcode, "error")
            continue

        if institutions is None:
            print(f"  NOT FOUND in Alma")
            results.append({
                "barcode": barcode,
                "status": "not_found",
                "title": None,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "error": f"Barcode not found via {leaving_school_name} API. "
                         f"The item may have been deleted, or the barcode/school code may be incorrect."
            })
            if progress_callback:
                progress_callback(i, total, barcode, "not_found")
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
            if progress_callback:
                progress_callback(i, total, barcode, "ineligible")
            continue

        title = bib_info.get("title", "Unknown") if bib_info else "Unknown"
        print(f"  Title: {title[:60]}...")
        print(f"  Held by: {len(institutions)} institution(s)")

        # Select replacement school
        replacement, eligible_list, no_replacement_reason = select_replacement_school(
            institutions, leaving_school, schools
        )

        if replacement is None:
            print(f"  ⚠️  NO ELIGIBLE REPLACEMENT - {no_replacement_reason}")
            results.append({
                "barcode": barcode,
                "status": "no_replacement",
                "title": title,
                "leaving_school": leaving_school,
                "replacement_school": None,
                "eligible_schools": [],
                "holding_institutions": institutions,
                "no_replacement_reason": no_replacement_reason,
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
                "holding_institutions": institutions,
                "bib_info": bib_info
            })

        # Notify caller of progress (used by Streamlit for the progress bar)
        if progress_callback:
            progress_callback(i, total, barcode, results[-1]["status"])

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


# =============================================================================
# PENDING TRANSFERS: FILE I/O
# =============================================================================

def save_pending_transfers(results, output_dir, schools):
    """
    Save the results of Phase 1 (lookup) to a JSON file so that Phase 2
    (update) can be run later, after email replies have been received.

    Each item is stored with:
      - All bib/item/holding IDs needed for Alma updates
      - The full eligible_schools priority list (so declined schools can be
        skipped and the next school tried automatically)
      - A declined_schools list (empty at first)
      - The proposed_school (first in the priority list)
      - status: "awaiting_reply" | "needs_review" | "no_replacement" |
                "not_found" | "ineligible" | "completed"
      - lookup_date: ISO timestamp of when the lookup was run

    Returns the path to the saved JSON file.
    """
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"pending_{timestamp}.json"
    filepath = os.path.join(output_dir, filename)

    # Build a serialisable version of results.
    # bib_info["item_data"] contains the raw Alma API response, which has
    # nested dicts/lists — all JSON-serialisable already.
    payload = {
        "version": 1,
        "created": datetime.now().isoformat(),
        "items": []
    }

    for r in results:
        # "replacement_found" means an email has been generated and sent;
        # in the pending file we call this "awaiting_reply" to reflect that
        # we are now waiting to hear back before making any record changes.
        raw_status = r.get("status")
        pending_status = "awaiting_reply" if raw_status == "replacement_found" else raw_status

        item_entry = {
            "barcode":           r.get("barcode"),
            "title":             r.get("title"),
            "status":            pending_status,
            "leaving_school":    r.get("leaving_school"),
            "proposed_school":   r.get("replacement_school"),   # first choice
            "eligible_schools":  r.get("eligible_schools", []),
            "declined_schools":  [],
            "lookup_date":       datetime.now().isoformat(),
            "bib_info":          r.get("bib_info"),             # may be None
        }
        # Carry through extra status fields where present
        for extra_key in ("item_status", "error"):
            if extra_key in r:
                item_entry[extra_key] = r[extra_key]
        payload["items"].append(item_entry)

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)

    return filepath


def load_pending_transfers(json_path):
    """
    Load a pending-transfers JSON file saved by save_pending_transfers().

    Returns the parsed payload dict, or exits with an error message if the
    file cannot be read or is not a valid pending-transfers file.
    """
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: Pending-transfers file not found: {json_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"ERROR: Could not parse pending-transfers file: {e}")
        sys.exit(1)

    if payload.get("version") != 1 or "items" not in payload:
        print(f"ERROR: '{json_path}' does not look like a pending-transfers file.")
        print("  Make sure you're passing the JSON file generated by the lookup step.")
        sys.exit(1)

    return payload


# =============================================================================
# STEP 1: LOOKUP PHASE  (Phases 1 + 2 — no record changes)
# =============================================================================

def run_lookup_phase(barcode_file, output_dir, config, schools):
    """
    Step 1: Look up barcodes, select replacement schools, generate draft
    emails, and save a pending-transfers JSON file.

    No Alma or WorldCat records are changed.
    """
    print("\n" + "=" * 60)
    print("STEP 1: LOOKUP")
    print("=" * 60)

    # Read barcodes
    items = read_barcodes(barcode_file)
    if not items:
        print("No items found. Exiting.")
        sys.exit(1)

    # Phase 1: Look up items and select replacement schools
    print("\nLooking up items and selecting replacement schools...")
    results = process_barcodes(items, schools, config)

    # Print summary
    print_summary(results, schools)

    # Phase 2: Generate draft emails
    email_dir = os.path.join(output_dir, "emails")
    print_draft_emails(results, schools, email_dir)

    # Save pending-transfers file
    pending_path = save_pending_transfers(results, output_dir, schools)
    print(f"\n✓ Pending-transfers file saved: {pending_path}")

    # Count actionable items
    awaiting = [r for r in results if r["status"] == "replacement_found"]
    if awaiting:
        print(f"\n{'=' * 60}")
        print("NEXT STEPS")
        print("=" * 60)
        print(f"  1. Open and send the draft email(s) in: {email_dir}/")
        print(f"  2. Wait for replies from the chief librarians.")
        print(f"  3. When you have a reply, run:")
        print(f"\n     python retention_transfer.py \\")
        print(f"         {pending_path} --update")
        if config.get("sandbox"):
            print(f"         --sandbox")
        print()
    else:
        print("\nNo actionable items found (no replacements to email).")
        print("Review the summary above for details.")

    return results, pending_path


# =============================================================================
# STEP 2: UPDATE PHASE  (Phases 3 + 4 + 5 — makes record changes)
# =============================================================================

def _re_verify_leaving_school_ids(result, schools, config):
    """
    Re-query Alma to confirm the leaving school's MMS ID, holding ID, and
    item PID are still valid.  Uses the barcode (which is stable) to look up
    the item fresh, then compares against the stored IDs.

    Returns (mms_id, holding_id, item_pid, warning_message)
    warning_message is None if everything matches, or a string describing any
    discrepancy.
    """
    barcode        = result["barcode"]
    leaving_code   = result["leaving_school"]
    stored_bib     = result.get("bib_info", {}) or {}
    stored_mms_id  = stored_bib.get("item_data", {}).get("bib_data", {}).get("mms_id", "")
    stored_holding = stored_bib.get("item_data", {}).get("holding_data", {}).get("holding_id", "")
    stored_pid     = stored_bib.get("item_data", {}).get("item_data", {}).get("pid", "")

    fresh = lookup_item_by_barcode(barcode, leaving_code, schools, config["base_url"])
    if not fresh:
        return None, None, None, "Could not re-verify leaving school item — barcode no longer found in Alma"

    fresh_mms_id  = fresh.get("bib_data",     {}).get("mms_id",     "")
    fresh_holding = fresh.get("holding_data",  {}).get("holding_id", "")
    fresh_pid     = fresh.get("item_data",     {}).get("pid",        "")

    warnings = []
    if stored_mms_id and fresh_mms_id != stored_mms_id:
        warnings.append(f"MMS ID changed: was {stored_mms_id}, now {fresh_mms_id}")
    if stored_holding and fresh_holding != stored_holding:
        warnings.append(f"Holding ID changed: was {stored_holding}, now {fresh_holding}")
    if stored_pid and fresh_pid != stored_pid:
        warnings.append(f"Item PID changed: was {stored_pid}, now {fresh_pid}")

    msg = "; ".join(warnings) if warnings else None
    return fresh_mms_id, fresh_holding, fresh_pid, msg


def _re_verify_taking_school_ids(result, taking_code, schools, config):
    """
    Re-query Alma to get fresh holding/item IDs for the taking school.

    Uses the NZ MMS ID (stored in bib_info) to look up the taking school's
    current IZ MMS ID, then retrieves holdings/items from that.

    Returns (iz_mms_id, holding_id, item_pid, error_message)
    error_message is None on success; begins with "NEEDS_REVIEW:" if ambiguous.
    """
    nz_mms_id = (result.get("bib_info") or {}).get("nz_mms_id")
    if not nz_mms_id:
        return None, None, None, "No NZ MMS ID stored — cannot re-verify taking school IDs"

    school   = schools.get(taking_code, {})
    api_key  = school.get("api_key", "")
    base_url = config["base_url"]

    iz_mms_id, err = get_taking_school_iz_mms_id(nz_mms_id, api_key, base_url)
    if err:
        return None, None, None, err

    holding_id, item_pid, err = get_taking_school_holding_and_item(iz_mms_id, api_key, base_url)
    if err:
        return None, None, None, err

    return iz_mms_id, holding_id, item_pid, None


def _handle_decline(item_entry, schools, email_dir):
    """
    Record a decline from the current proposed_school, advance to the next
    eligible school, and generate a new draft email if one is available.

    Mutates item_entry in place.
    Returns a human-readable message describing what happened.
    """
    declined = item_entry["proposed_school"]
    if declined and declined not in item_entry["declined_schools"]:
        item_entry["declined_schools"].append(declined)

    # Find the next eligible school not yet declined
    remaining = [
        s for s in item_entry["eligible_schools"]
        if s not in item_entry["declined_schools"]
    ]

    if not remaining:
        item_entry["proposed_school"] = None
        item_entry["status"] = "no_replacement"
        return "All eligible schools have declined. Flagged for withdrawal review."

    next_school = remaining[0]
    item_entry["proposed_school"] = next_school
    item_entry["status"] = "awaiting_reply"

    # Build a minimal result dict that generate_draft_email expects
    fake_result = {
        "barcode":           item_entry["barcode"],
        "title":             item_entry["title"],
        "status":            "replacement_found",
        "replacement_school": next_school,
        "bib_info":          item_entry.get("bib_info"),
    }

    if email_dir:
        os.makedirs(email_dir, exist_ok=True)
        email = generate_draft_email([fake_result], next_school, schools)
        filepath = save_eml_file(email, email_dir)
        return f"Moved to next school: {schools[next_school]['name']}. New email: {filepath}"
    else:
        return f"Moved to next school: {schools[next_school]['name']}."


def run_update_phase(json_path, output_dir, config, schools):
    """
    Step 2: Process replies to the emails sent in Step 1.

    For each item that is 'awaiting_reply', prompts:
        Did [school name] agree? (yes / no / skip)

    yes   — re-queries Alma to get fresh IDs, then runs Phases 3, 4, 5
    no    — records the decline, moves to next eligible school, new email
    skip  — leaves the item pending for the next run

    The pending JSON file is updated at the end of every run, so this
    command can be re-run as many times as needed.
    """
    payload = load_pending_transfers(json_path)
    items   = payload["items"]
    created = payload.get("created", "unknown date")

    email_dir      = os.path.join(output_dir, "emails")
    worldcat_dir   = output_dir   # Phase 5 derives subdirs from this

    print("\n" + "=" * 60)
    print("STEP 2: UPDATE")
    print("=" * 60)
    print(f"  Pending file: {json_path}")
    print(f"  Lookup date:  {created}")

    # Tally statuses
    awaiting   = [it for it in items if it["status"] == "awaiting_reply"]
    completed  = [it for it in items if it["status"] == "completed"]
    no_replace = [it for it in items if it["status"] == "no_replacement"]
    other      = [it for it in items if it["status"] not in
                  ("awaiting_reply", "completed", "no_replacement")]

    print(f"\n  Awaiting reply:   {len(awaiting)}")
    print(f"  Already complete: {len(completed)}")
    print(f"  No replacement:   {len(no_replace)}")
    if other:
        print(f"  Other (skipped):  {len(other)}")

    if not awaiting:
        print("\nNo items awaiting a reply. Nothing to do.")
        _print_update_summary(items, schools)
        return

    if config.get("sandbox"):
        print("\n  ⚠️  SANDBOX MODE — changes will be made to sandbox only")
    else:
        print("\n  ⚠️  PRODUCTION MODE — changes will be made to live Alma records")

    # -------------------------------------------------------------------------
    # Per-item confirmation loop
    # -------------------------------------------------------------------------
    newly_completed = []
    newly_declined  = []
    skipped         = []
    errors          = []

    for i, item_entry in enumerate(awaiting, 1):
        title        = item_entry.get("title", "Unknown")
        barcode      = item_entry.get("barcode", "")
        leaving_code = item_entry.get("leaving_school", "")
        taking_code  = item_entry.get("proposed_school", "")
        lookup_date  = item_entry.get("lookup_date", created)

        leaving_name = schools.get(leaving_code, {}).get("name", leaving_code)
        taking_name  = schools.get(taking_code,  {}).get("name", taking_code) if taking_code else "None"
        declined_names = [schools.get(s, {}).get("name", s)
                          for s in item_entry.get("declined_schools", [])]

        print(f"\n{'─' * 60}")
        print(f"[{i}/{len(awaiting)}] {title[:70]}")
        print(f"  Barcode:       {barcode}")
        print(f"  Leaving:       {leaving_name}")
        print(f"  Proposed:      {taking_name}")
        if declined_names:
            print(f"  Declined so far: {', '.join(declined_names)}")
        print(f"  Lookup date:   {lookup_date[:10]}")

        if not taking_code:
            print("  ⚠️  No proposed school — skipping (already exhausted all options?)")
            skipped.append(barcode)
            continue

        answer = ""
        while answer not in ("yes", "no", "skip"):
            answer = input(f"\n  Did {taking_name} agree to take this commitment? (yes/no/skip): ").strip().lower()
            if answer not in ("yes", "no", "skip"):
                print("  Please type 'yes', 'no', or 'skip'.")

        if answer == "skip":
            print(f"  Skipping — will try again next run.")
            skipped.append(barcode)
            continue

        if answer == "no":
            msg = _handle_decline(item_entry, schools, email_dir)
            print(f"  {msg}")
            newly_declined.append(barcode)
            continue

        # ── answer == "yes" ──────────────────────────────────────────────────
        print(f"\n  {taking_name} agreed. Re-verifying Alma IDs before making changes...")

        # Re-verify leaving school IDs
        l_mms_id, l_holding_id, l_item_pid, warn = _re_verify_leaving_school_ids(
            item_entry, schools, config
        )
        if l_mms_id is None:
            print(f"  ✗ Could not verify leaving school IDs: {warn}")
            print(f"    Skipping this item — no changes made.")
            errors.append({"barcode": barcode, "message": warn})
            continue
        if warn:
            print(f"  ⚠️  Leaving school IDs changed since lookup: {warn}")
            print(f"    Using updated IDs.")

        # Re-verify taking school IDs
        t_iz_mms_id, t_holding_id, t_item_pid, err = _re_verify_taking_school_ids(
            item_entry, taking_code, schools, config
        )
        if err:
            if err.startswith("NEEDS_REVIEW:"):
                print(f"  ⚠️  Taking school needs manual review: {err[len('NEEDS_REVIEW:'):].strip()}")
            else:
                print(f"  ✗ Could not verify taking school IDs: {err}")
            print(f"    Skipping this item — no changes made.")
            errors.append({"barcode": barcode, "message": err})
            continue

        print(f"  IDs verified. Proceeding with updates...")

        # Phase 3: Update leaving school
        leaving_school  = schools[leaving_code]
        l_api_key       = leaving_school.get("api_key", "")
        item_ok, item_msg = update_leaving_school_item(
            l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
        )
        if item_ok:
            print(f"  ✓ Leaving school item: {item_msg}")
        else:
            print(f"  ✗ Leaving school item update failed: {item_msg}")
            errors.append({"barcode": barcode, "message": item_msg})
            continue

        holdings_ok, holdings_msg = update_leaving_school_holdings(
            l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
        )
        if holdings_ok:
            print(f"  ✓ Leaving school holdings: {holdings_msg}")
        else:
            print(f"  ✗ Leaving school holdings update failed: {holdings_msg}")
            # Non-fatal: continue to Phase 4

        # Phase 4: Update taking school
        taking_school   = schools[taking_code]
        t_api_key       = taking_school.get("api_key", "")
        marc_org_code   = taking_school.get("marc_org_code", "")
        try:
            marc_org_missing = not marc_org_code or pd.isna(marc_org_code) or str(marc_org_code).lower() == "nan"
        except (TypeError, ValueError):
            marc_org_missing = True
        if marc_org_missing:
            marc_org_code = ""

        item_ok, item_msg = update_taking_school_item(
            t_iz_mms_id, t_holding_id, t_item_pid, t_api_key, config["base_url"]
        )
        if item_ok:
            print(f"  ✓ Taking school item: {item_msg}")
        else:
            print(f"  ✗ Taking school item update failed: {item_msg}")
            errors.append({"barcode": barcode, "message": item_msg})
            continue

        holdings_ok, holdings_msg = update_taking_school_holdings(
            t_iz_mms_id, t_holding_id, marc_org_code, t_api_key, config["base_url"]
        )
        if holdings_ok:
            print(f"  ✓ Taking school holdings: {holdings_msg}")
        else:
            print(f"  ✗ Taking school holdings update failed: {holdings_msg}")
            # Non-fatal: mark completed anyway, note the issue

        # Phase 5: WorldCat CSV for this item
        # Build a minimal result dict matching what generate_worldcat_taking_csv expects
        worldcat_result = {
            "status":             "replacement_found",
            "barcode":            barcode,
            "title":              title,
            "replacement_school": taking_code,
            "leaving_school":     leaving_code,
            "bib_info":           item_entry.get("bib_info"),
        }
        wc_files, wc_skipped = generate_worldcat_taking_csv(
            [worldcat_result], schools, worldcat_dir
        )
        if wc_files:
            print(f"  ✓ WorldCat CSV: {wc_files[0]}")
        elif wc_skipped:
            print(f"  ⚠️  WorldCat CSV skipped: {wc_skipped[0]['reason']}")

        wc_instructions = generate_worldcat_leaving_instructions(
            [worldcat_result], schools, worldcat_dir
        )
        if wc_instructions:
            print(f"  ✓ WorldCat leaving instructions: {wc_instructions}")

        # Mark completed
        item_entry["status"]         = "completed"
        item_entry["completed_date"] = datetime.now().isoformat()
        item_entry["taking_school"]  = taking_code
        newly_completed.append(barcode)
        print(f"  ✓ Transfer complete.")

    # -------------------------------------------------------------------------
    # No-replacement items: offer to remove the leaving school's commitment
    # -------------------------------------------------------------------------
    no_replacement_items = [it for it in items if it["status"] == "no_replacement"]
    if no_replacement_items:
        print(f"\n{'─' * 60}")
        print(f"NO-REPLACEMENT ITEMS ({len(no_replacement_items)})")
        print(f"{'─' * 60}")
        print("The following items have no eligible replacement school.")
        print("The leaving school's retention commitment can still be removed")
        print("even though no other school is taking it on.\n")
        for it in no_replacement_items:
            leaving_code = it.get("leaving_school", "")
            leaving_name = schools.get(leaving_code, {}).get("name", leaving_code)
            print(f"  {it['barcode']}  {(it.get('title') or '')[:50]}")
            print(f"  Leaving: {leaving_name}")

        answer = ""
        while answer not in ("yes", "no"):
            answer = input(
                f"\nRemove retention commitments from the leaving school's "
                f"Alma records for these {len(no_replacement_items)} item(s)? (yes/no): "
            ).strip().lower()
            if answer not in ("yes", "no"):
                print("  Please type 'yes' or 'no'.")

        if answer == "yes":
            for item_entry in no_replacement_items:
                barcode      = item_entry.get("barcode", "?")
                leaving_code = item_entry.get("leaving_school", "")
                title        = item_entry.get("title", "Unknown")

                print(f"\n  Processing {barcode} — {title[:50]}")

                # Re-verify leaving school IDs before making changes
                l_mms_id, l_holding_id, l_item_pid, warn = _re_verify_leaving_school_ids(
                    item_entry, schools, config
                )
                if l_mms_id is None:
                    print(f"  ✗ Could not verify leaving school IDs: {warn}")
                    errors.append({"barcode": barcode, "message": warn})
                    continue
                if warn:
                    print(f"  ⚠️  IDs changed since lookup: {warn}")

                leaving_school = schools.get(leaving_code, {})
                l_api_key      = leaving_school.get("api_key", "")

                item_ok, item_msg = update_leaving_school_item(
                    l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
                )
                if item_ok:
                    print(f"  ✓ Leaving school item: {item_msg}")
                else:
                    print(f"  ✗ Leaving school item update failed: {item_msg}")
                    errors.append({"barcode": barcode, "message": item_msg})
                    continue

                _, holdings_msg = update_leaving_school_holdings(
                    l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
                )
                print(f"  ✓ Leaving school holdings: {holdings_msg}")

                # WorldCat leaving instructions (no taking school CSV)
                worldcat_result = {
                    "status":             "replacement_found",
                    "barcode":            barcode,
                    "title":              title,
                    "replacement_school": None,
                    "leaving_school":     leaving_code,
                    "bib_info":           item_entry.get("bib_info"),
                }
                wc_instructions = generate_worldcat_leaving_instructions(
                    [worldcat_result], schools, worldcat_dir
                )
                if wc_instructions:
                    print(f"  ✓ WorldCat leaving instructions: {wc_instructions}")

                item_entry["status"]         = "commitment_removed"
                item_entry["completed_date"] = datetime.now().isoformat()
                print(f"  ✓ Commitment removed.")
        else:
            print("  Skipping — commitments left in place.")

    # -------------------------------------------------------------------------
    # Save updated pending file
    # -------------------------------------------------------------------------
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)
    print(f"\n✓ Pending file updated: {json_path}")

    # -------------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------------
    _print_update_summary(items, schools)

    # Prompt for next steps if there are new emails to send
    still_awaiting = [it for it in items if it["status"] == "awaiting_reply"]
    if still_awaiting:
        print(f"\n  {len(still_awaiting)} item(s) still awaiting a reply.")
        print(f"  New draft emails (if any) were saved to: {email_dir}/")
        print(f"  When you have more replies, run:")
        print(f"\n     python retention_transfer.py \\")
        print(f"         {json_path} --update")
        if config.get("sandbox"):
            print(f"         --sandbox")

    if errors:
        print(f"\n  ⚠️  {len(errors)} item(s) had errors and were NOT updated:")
        for e in errors:
            print(f"    - {e['barcode']}: {e['message']}")


def _print_update_summary(items, schools):
    """Print a tally of all item statuses in the pending file."""
    print(f"\n{'=' * 60}")
    print("STATUS OF ALL ITEMS")
    print("=" * 60)

    by_status = {}
    for it in items:
        s = it["status"]
        by_status.setdefault(s, []).append(it)

    status_labels = {
        "completed":           "✓ Completed (transferred)",
        "commitment_removed":  "✓ Commitment removed (no replacement)",
        "awaiting_reply":      "⏳ Awaiting reply",
        "no_replacement":      "⚠️  No replacement (withdrawal review)",
        "not_found":           "✗ Not found in Alma",
        "ineligible":          "✗ Ineligible (item not in place)",
        "error":               "✗ Error",
    }
    for status, label in status_labels.items():
        group = by_status.get(status, [])
        if group:
            print(f"\n  {label} ({len(group)}):")
            for it in group:
                taking = it.get("taking_school") or it.get("proposed_school")
                taking_name = schools.get(taking, {}).get("name", taking) if taking else "—"
                leaving_name = schools.get(it.get("leaving_school", ""), {}).get("name", "—")
                print(f"    {it['barcode']}  {it.get('title', '')[:45]}")
                print(f"      From: {leaving_name}  →  To: {taking_name}")

    print("=" * 60)


# =============================================================================
# MAIN
# =============================================================================

def main():
    """
    Entry point. Dispatches to lookup mode or update mode based on arguments.

    Lookup mode (default):
        python retention_transfer.py data/barcodes.xlsx [--sandbox]

    Update mode:
        python retention_transfer.py output/pending_YYYYMMDD_HHMMSS.json --update [--sandbox]
    """
    flags    = [a for a in sys.argv[1:] if a.startswith("--")]
    args     = [a for a in sys.argv[1:] if not a.startswith("--")]
    sandbox  = "--sandbox" in flags
    update   = "--update"  in flags

    if len(args) < 1:
        print(__doc__)
        sys.exit(1)

    input_file = args[0]
    output_dir = args[1] if len(args) > 1 else "output"

    print("=" * 60)
    print("CUNY Shared Print - Retention Transfer")
    print("=" * 60)

    config = get_config(sandbox=sandbox)
    if sandbox:
        print("⚠️  SANDBOX MODE")
    print(f"API Base URL: {config['base_url']}")

    schools = load_schools(config["schools_file"])
    print(f"Loaded {len(schools)} schools")

    if update:
        # ── UPDATE MODE: first arg must be a .json pending-transfers file ──
        if not input_file.endswith(".json"):
            print(f"ERROR: --update requires a pending-transfers JSON file as the first argument.")
            print(f"  Got: {input_file}")
            print(f"  Example: python retention_transfer.py output/pending_20260220_143012.json --update")
            sys.exit(1)
        run_update_phase(input_file, output_dir, config, schools)
    else:
        # ── LOOKUP MODE: first arg must be an Excel file ──
        if not (input_file.endswith(".xlsx") or input_file.endswith(".xls")):
            print(f"ERROR: Lookup mode requires an Excel (.xlsx) file as the first argument.")
            print(f"  Got: {input_file}")
            print(f"  To run updates on an existing pending file, add --update.")
            sys.exit(1)
        run_lookup_phase(input_file, output_dir, config, schools)


if __name__ == "__main__":
    main()
