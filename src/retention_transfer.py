"""
CUNY Shared Print Monograph Trust - Retention Transfer Script

This script helps transfer retention commitments from one school to another
when a school can no longer retain a book.

Phase 1: Lookup and school selection
- Reads barcodes and school codes from Excel file
- Looks up each item using the specified school's API key
- Queries Network Zone to find all schools holding the title
- Selects replacement school based on rules

Usage:
    python retention_transfer.py barcodes.xlsx

The Excel file should have two columns:
    - Barcode: The item barcode
    - School Code: The Alma Institution Code (e.g., 01CUNY_QC)

Example:
    python retention_transfer.py data/barcodes.xlsx
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

def get_config():
    """Load configuration from environment variables."""
    config = {
        "nz_api_key": os.getenv("ALMA_NZ_API_KEY"),  # Network Zone key
        "base_url": os.getenv("ALMA_API_BASE_URL", "https://api-na.hosted.exlibrisgroup.com"),
        "schools_file": os.getenv("SCHOOLS_FILE", "data/schools_template.csv")
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
            "chief_librarian_name": row.get("Chief Librarian Name", ""),
            "chief_librarian_email": row.get("Chief Librarian Email", ""),
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


def find_holding_institutions(barcode, leaving_school, schools, config):
    """
    Find all institutions that hold a copy of this item.

    1. Look up the barcode using the specified school's API key
    2. Extract the Network Zone MMS ID
    3. Query the Network Zone to find all holding institutions

    Args:
        barcode: The item barcode
        leaving_school: The Alma Institution Code of the school leaving retention
        schools: Dictionary of school data
        config: Configuration dictionary

    Returns: (list of institution codes, bib_info dict)
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

    print(f"\nReplacement found:     {len(found)}")
    print(f"No replacement (review): {len(no_replacement)}")
    print(f"Not found in Alma:     {len(not_found)}")

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

    if not_found:
        print("\n--- Barcodes Not Found ---")
        for r in not_found:
            print(f"  {r['barcode']}")

    print("=" * 60)

    return found, no_replacement, not_found


def main():
    """Main function."""
    # Check arguments
    if len(sys.argv) < 2:
        print("Usage: python retention_transfer.py <barcodes.xlsx>")
        print("\nThe Excel file should have two columns:")
        print("  - Barcode: The item barcode")
        print("  - School Code: The Alma Institution Code (e.g., 01CUNY_QC)")
        print("\nExample:")
        print("  python retention_transfer.py data/barcodes.xlsx")
        sys.exit(1)

    barcode_file = sys.argv[1]

    print("=" * 60)
    print("CUNY Shared Print - Retention Transfer (Phase 1)")
    print("=" * 60)

    # Load configuration
    config = get_config()
    print(f"API Base URL: {config['base_url']}")

    # Load schools data
    schools = load_schools(config["schools_file"])
    print(f"Loaded {len(schools)} schools")

    # Read barcodes and school codes
    items = read_barcodes(barcode_file)

    if not items:
        print("No items found. Exiting.")
        sys.exit(1)

    # Process barcodes
    print("\nLooking up items and selecting replacement schools...")
    results = process_barcodes(items, schools, config)

    # Print summary
    found, no_replacement, not_found = print_summary(results, schools)

    # TODO Phase 2: Generate draft emails for 'found' items
    # TODO Phase 3-5: Update records after confirmation

    print("\nPhase 1 complete. Next steps:")
    print("- Review the replacement selections above")
    print("- Phase 2 will generate draft emails to send to replacement schools")


if __name__ == "__main__":
    main()
