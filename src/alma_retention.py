"""
Library Retention Records - Alma Update Script

This script reads barcodes from an Excel file and updates the
retention reason field for matching items in Alma.

Usage:
    python alma_retention.py path/to/barcodes.xlsx

Requirements:
    - Alma API key with Bibs read/write permissions
    - Excel file with barcodes in a column
"""

import os
import sys
import requests
import pandas as pd
from dotenv import load_dotenv

# Load API key from .env file
load_dotenv()


def get_api_key():
    """Get the Alma API key from environment variables."""
    api_key = os.getenv("ALMA_API_KEY")
    if not api_key:
        print("ERROR: No API key found!")
        print("Make sure you have a .env file with ALMA_API_KEY=your_key_here")
        sys.exit(1)
    return api_key


def get_alma_base_url():
    """Get the Alma API base URL from environment variables."""
    # Default to North America region - change if your institution is elsewhere
    base_url = os.getenv("ALMA_API_BASE_URL", "https://api-na.hosted.exlibrisgroup.com")
    return base_url


def read_barcodes_from_excel(file_path, column_name=None):
    """
    Read barcodes from an Excel file.

    Args:
        file_path: Path to the Excel file
        column_name: Name of the column containing barcodes (optional)
                    If not provided, uses the first column

    Returns:
        List of barcodes as strings
    """
    print(f"Reading barcodes from: {file_path}")

    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Could not read Excel file: {e}")
        sys.exit(1)

    # If no column specified, use the first column
    if column_name is None:
        column_name = df.columns[0]
        print(f"Using first column: '{column_name}'")

    # Get barcodes and convert to strings (in case they're numbers)
    barcodes = df[column_name].astype(str).tolist()

    # Remove any empty or 'nan' values
    barcodes = [b for b in barcodes if b and b.lower() != 'nan']

    print(f"Found {len(barcodes)} barcodes")
    return barcodes


def lookup_item_by_barcode(barcode, api_key, base_url):
    """
    Look up an item in Alma by its barcode.

    Args:
        barcode: The item barcode to search for
        api_key: Alma API key
        base_url: Alma API base URL

    Returns:
        Dictionary with item data, or None if not found
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
            print(f"  Barcode {barcode}: Not found in Alma")
            return None
        else:
            print(f"  Barcode {barcode}: API error (status {response.status_code})")
            return None

    except requests.RequestException as e:
        print(f"  Barcode {barcode}: Connection error - {e}")
        return None


def update_item_retention(item_data, retention_reason, api_key, base_url):
    """
    Update an item's retention reason field in Alma.

    Args:
        item_data: The item data returned from lookup
        retention_reason: The value to set for retention reason
        api_key: Alma API key
        base_url: Alma API base URL

    Returns:
        True if successful, False otherwise
    """
    # Extract the IDs needed for the update URL
    try:
        mms_id = item_data["bib_data"]["mms_id"]
        holding_id = item_data["holding_data"]["holding_id"]
        item_pid = item_data["item_data"]["pid"]
        barcode = item_data["item_data"]["barcode"]
    except KeyError as e:
        print(f"  ERROR: Missing expected field in item data: {e}")
        return False

    # Build the update URL
    url = f"{base_url}/almaws/v1/bibs/{mms_id}/holdings/{holding_id}/items/{item_pid}"

    params = {
        "apikey": api_key
    }

    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    # Update the retention reason in the item data
    item_data["item_data"]["retention_reason"] = {
        "value": retention_reason
    }

    try:
        response = requests.put(url, params=params, headers=headers, json=item_data)

        if response.status_code == 200:
            print(f"  Barcode {barcode}: Successfully updated retention reason")
            return True
        else:
            print(f"  Barcode {barcode}: Update failed (status {response.status_code})")
            # Try to show error details
            try:
                error_info = response.json()
                if "errorList" in error_info:
                    for error in error_info["errorList"]["error"]:
                        print(f"    - {error.get('errorMessage', 'Unknown error')}")
            except Exception:
                pass
            return False

    except requests.RequestException as e:
        print(f"  Barcode {barcode}: Connection error - {e}")
        return False


def process_barcodes(barcodes, retention_reason, api_key, base_url):
    """
    Process a list of barcodes: look up each one and update retention.

    Args:
        barcodes: List of barcodes to process
        retention_reason: The retention reason value to set
        api_key: Alma API key
        base_url: Alma API base URL

    Returns:
        Dictionary with counts of successes, failures, and not found
    """
    results = {
        "success": 0,
        "failed": 0,
        "not_found": 0
    }

    total = len(barcodes)

    for i, barcode in enumerate(barcodes, 1):
        print(f"\nProcessing {i}/{total}: {barcode}")

        # Look up the item
        item_data = lookup_item_by_barcode(barcode, api_key, base_url)

        if item_data is None:
            results["not_found"] += 1
            continue

        # Update the retention reason
        success = update_item_retention(item_data, retention_reason, api_key, base_url)

        if success:
            results["success"] += 1
        else:
            results["failed"] += 1

    return results


def main():
    """Main function to run the script."""

    # Check command line arguments
    if len(sys.argv) < 2:
        print("Usage: python alma_retention.py path/to/barcodes.xlsx [retention_reason]")
        print("\nExample:")
        print("  python alma_retention.py barcodes.xlsx 'Shared Print Retention'")
        sys.exit(1)

    excel_file = sys.argv[1]

    # Get retention reason (default or from command line)
    retention_reason = "Committed to Retain"
    if len(sys.argv) >= 3:
        retention_reason = sys.argv[2]

    print("=" * 50)
    print("Library Retention Records - Alma Update Script")
    print("=" * 50)

    # Get API credentials
    api_key = get_api_key()
    base_url = get_alma_base_url()

    print(f"API Base URL: {base_url}")
    print(f"Retention reason: {retention_reason}")

    # Read barcodes from Excel
    barcodes = read_barcodes_from_excel(excel_file)

    if not barcodes:
        print("No barcodes found in file. Exiting.")
        sys.exit(1)

    # Confirm before proceeding
    print(f"\nReady to update {len(barcodes)} items in Alma.")
    response = input("Continue? (yes/no): ")

    if response.lower() not in ["yes", "y"]:
        print("Cancelled.")
        sys.exit(0)

    # Process the barcodes
    print("\nStarting updates...")
    results = process_barcodes(barcodes, retention_reason, api_key, base_url)

    # Print summary
    print("\n" + "=" * 50)
    print("SUMMARY")
    print("=" * 50)
    print(f"Successfully updated: {results['success']}")
    print(f"Not found in Alma:    {results['not_found']}")
    print(f"Failed to update:     {results['failed']}")
    print("=" * 50)


if __name__ == "__main__":
    main()
