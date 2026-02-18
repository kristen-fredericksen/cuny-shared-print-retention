# CUNY Shared Print Monograph Trust - Retention Transfer

A Python script to manage retention commitment transfers between CUNY libraries participating in the Shared Print program.

## What This Does

When a CUNY library can no longer retain a book it committed to keep, this script helps transfer that commitment to another library:

1. **Looks up the item** in Alma using the leaving school's API
2. **Finds all CUNY schools** that hold a copy (via Network Zone)
3. **Selects a replacement school** using these rules:
   - Must be a Shared Print participant
   - Prefers CUNY Graduate Center if they hold it
   - Otherwise picks the largest participating school that holds it
4. **Flags items for withdrawal review** if no other schools hold them

## Project Status

| Phase | Description | Status |
|-------|-------------|--------|
| **Phase 1** | Barcode lookup and replacement selection | ✅ Complete |
| **Phase 2** | Draft email generation | ✅ Complete |
| **Phase 3** | Update leaving school's Alma records | 🔜 Planned |
| **Phase 4** | Update taking school's Alma records | 🔜 Planned |
| **Phase 5** | Update WorldCat holdings | 🔜 Planned |

## Setup

### Prerequisites
- Python 3.x
- Alma API keys (Network Zone + Institution Zone keys for participating schools)

### Installation

```bash
# Clone the repository
git clone https://github.com/kristen-fredericksen/cuny-shared-print-retention.git
cd cuny-shared-print-retention

# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Set up configuration
cp .env.example .env
# Edit .env with your API keys
```

### Configuration

1. **`.env` file** - Contains your Network Zone API key:
   ```
   ALMA_NZ_API_KEY=your_network_zone_key
   ALMA_API_BASE_URL=https://api-na.hosted.exlibrisgroup.com
   SCHOOLS_FILE=data/schools_template.csv
   ```

2. **`data/schools_template.csv`** - Contains school information:
   - Name, Size (1=largest), Shared Print (Yes/No)
   - Alma Institution Code, OCLC Symbol, Primo View ID
   - Chief Librarian Name and Email
   - Alma API Key (for each school)

## Usage

### Input File Format

Create an Excel file with two columns:

| Barcode | School Code |
|---------|-------------|
| 31699001195116 | 01CUNY_JJ |
| 31228005903455 | 01CUNY_BC |

- **Barcode**: The item's barcode
- **School Code**: The Alma Institution Code of the school leaving retention

### Running the Script

```bash
# Run with default output directory (output/emails)
python src/retention_transfer.py data/your_barcodes.xlsx

# Or specify a custom output directory
python src/retention_transfer.py data/your_barcodes.xlsx output/emails
```

### Output

The script will:
1. Display a summary of items:
   - Items with recommended replacement schools
   - Items flagged for withdrawal review (no other schools hold them)
   - Items ineligible due to status (not "Item in place")
   - Barcodes not found in Alma

2. Generate `.eml` files for each replacement school:
   - Saved to `output/emails/` by default
   - Double-click to open in Outlook as a draft message
   - Includes school-specific Primo VE links for each title
   - Batches multiple titles per school into one email

## Record Updates (Phases 3-5)

When updating records, the following changes will be made:

**For the LEAVING school:**
- Alma holdings record: Remove 583 field
- Alma item record: Set "Committed to Retain" to "No"
- Alma item record: Clear "Retention Reason"
- WorldCat LHR: Remove 583 field

**For the TAKING school:**
- Alma holdings record: Add 583 field
- Alma item record: Set "Committed to Retain" to "Yes"
- Alma item record: Set "Retention Reason" to "CUNY Shared Print"
- WorldCat LHR: Add 583 field

## Files Not in Repository

The following files contain sensitive data and are excluded from git:
- `.env` (API keys)
- `data/` folder (school information and barcode files)

## License

Internal CUNY use only.
