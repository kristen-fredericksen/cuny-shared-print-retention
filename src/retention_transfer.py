"""
CUNY Shared Print Monograph Trust - Retention Transfer Script

This script helps transfer retention commitments from one school to another
when a school can no longer retain a book.

Phase 1: Lookup and school selection
- Reads barcodes from Excel file
- Looks up holdings in Alma Network Zone
- Selects replacement school based on rules

Usage:
    python retention_transfer.py barcodes.xlsx leaving_school_code

Example:
    python retention_transfer.py data/barcodes.xlsx 01CUNY_BC
"""

import os
import sys
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
        "api_key": os.getenv("ALMA_API_KEY"),
        "base_url": os.getenv("ALMA_API_BASE_URL", "https://api-na.hosted.exlibrisgroup.com"),
        "schools_file": os.getenv("SCHOOLS_FILE", "data/schools.csv")
    }

    if not config["api_key"]:
        print("ERROR: No API key found!")
        print("Make sure you have a .env file with ALMA_API_KEY=your_key_here")
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

def lookup_item_by_barcode(barcode, api_key, base_url):
    """
    Look up an item in Alma by its barcode.
    Returns the item data including which institution holds it.
    """
    url = f"{base_url}/almaws/v1/items"

    params = {
        "item_barcode": barcode,
        "apikey": api_key
    }

    headers = {
        "Accept": "application/json"
    }

    try:
        response = requests.get(url, params=params, headers=headers)

        if response.status_code == 200:
            return response.json()
        elif response.status_code == 400:
            return None  # Not found
        else:
            print(f"  API error (status {response.status_code})")
            return None

    except requests.RequestException as e:
        print(f"  Connection error: {e}")
        return None


def get_network_zone_holdings(mms_id, api_key, base_url):
    """
    Get all holdings for a bibliographic record from the Network Zone.
    This shows which institutions have holdings linked to this record.

    Note: This requires Network Zone API access.
    """
    url = f"{base_url}/almaws/v1/bibs/{mms_id}/holdings"

    params = {
        "apikey": api_key
    }

    headers = {
        "Accept": "application/json"
    }

    try:
        response = requests.get(url, params=params, headers=headers)

        if response.status_code == 200:
            return response.json()
        else:
            print(f"  Could not get holdings (status {response.status_code})")
            return None

    except requests.RequestException as e:
        print(f"  Connection error: {e}")
        return None


def find_holding_institutions(barcode, api_key, base_url):
    """
    Find all institutions that hold a copy of this item.

    Returns a list of institution codes.
    """
    # First, look up the item to get its MMS ID
    item_data = lookup_item_by_barcode(barcode, api_key, base_url)

    if not item_data:
        return None, None

    # Get the MMS ID (bibliographic record ID)
    try:
        mms_id = item_data["bib_data"]["mms_id"]
        title = item_data["bib_data"].get("title", "Unknown title")
    except KeyError:
        print("  Could not extract MMS ID from item data")
        return None, None

    # Get all holdings for this bib record
    holdings_data = get_network_zone_holdings(mms_id, api_key, base_url)

    if not holdings_data or "holding" not in holdings_data:
        return [], {"mms_id": mms_id, "title": title, "item_data": item_data}

    # Extract institution codes from holdings
    institutions = []
    for holding in holdings_data["holding"]:
        # The institution code might be in different places depending on setup
        inst_code = holding.get("institution", {}).get("value", "")
        if inst_code:
            institutions.append(inst_code)

    return institutions, {"mms_id": mms_id, "title": title, "item_data": item_data}


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
    """Read barcodes from Excel file."""
    print(f"Reading barcodes from: {file_path}")

    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    # Use first column
    column = df.columns[0]
    barcodes = df[column].astype(str).tolist()
    barcodes = [b for b in barcodes if b and b.lower() != 'nan']

    print(f"Found {len(barcodes)} barcodes")
    return barcodes


def process_barcodes(barcodes, leaving_school, schools, config):
    """
    Process each barcode: look up holdings and select replacement school.

    Returns a list of results for each barcode.
    """
    results = []
    total = len(barcodes)

    for i, barcode in enumerate(barcodes, 1):
        print(f"\n[{i}/{total}] Processing barcode: {barcode}")

        # Find which institutions hold this item
        institutions, bib_info = find_holding_institutions(
            barcode, config["api_key"], config["base_url"]
        )

        if institutions is None:
            print(f"  NOT FOUND in Alma")
            results.append({
                "barcode": barcode,
                "status": "not_found",
                "title": None,
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
            school_name = schools[r["replacement_school"]]["name"]
            print(f"  {r['barcode']}: {school_name}")
            print(f"    Title: {r['title'][:50]}...")

    if no_replacement:
        print("\n--- Items Flagged for Withdrawal Review ---")
        for r in no_replacement:
            print(f"  {r['barcode']}: {r['title'][:50]}...")

    if not_found:
        print("\n--- Barcodes Not Found ---")
        for r in not_found:
            print(f"  {r['barcode']}")

    print("=" * 60)

    return found, no_replacement, not_found


def main():
    """Main function."""
    # Check arguments
    if len(sys.argv) < 3:
        print("Usage: python retention_transfer.py <barcodes.xlsx> <leaving_school_code>")
        print("\nExample:")
        print("  python retention_transfer.py data/barcodes.xlsx 01CUNY_BC")
        print("\nThe leaving_school_code is the Alma Institution Code of the school")
        print("that can no longer retain the items.")
        sys.exit(1)

    barcode_file = sys.argv[1]
    leaving_school = sys.argv[2]

    print("=" * 60)
    print("CUNY Shared Print - Retention Transfer (Phase 1)")
    print("=" * 60)

    # Load configuration
    config = get_config()
    print(f"API Base URL: {config['base_url']}")

    # Load schools data
    schools = load_schools(config["schools_file"])
    print(f"Loaded {len(schools)} schools")

    # Verify leaving school exists
    if leaving_school not in schools:
        print(f"\nERROR: School code '{leaving_school}' not found in schools file.")
        print("Available codes:", list(schools.keys()))
        sys.exit(1)

    print(f"Leaving school: {schools[leaving_school]['name']} ({leaving_school})")

    # Read barcodes
    barcodes = read_barcodes(barcode_file)

    if not barcodes:
        print("No barcodes found. Exiting.")
        sys.exit(1)

    # Process barcodes
    print("\nLooking up items and selecting replacement schools...")
    results = process_barcodes(barcodes, leaving_school, schools, config)

    # Print summary
    found, no_replacement, not_found = print_summary(results, schools)

    # TODO Phase 2: Generate draft emails for 'found' items
    # TODO Phase 3-5: Update records after confirmation

    print("\nPhase 1 complete. Next steps:")
    print("- Review the replacement selections above")
    print("- Phase 2 will generate draft emails to send to replacement schools")


if __name__ == "__main__":
    main()
