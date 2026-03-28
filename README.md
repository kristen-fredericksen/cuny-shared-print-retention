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
| **Phase 3** | Update leaving school's Alma records | ✅ Complete |
| **Phase 4** | Update taking school's Alma records | ✅ Complete |
| **Phase 5** | Update WorldCat holdings | ✅ Complete |

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

1. **`.env` file** - Contains all API keys. See `.env.example` for the full format:
   - Network Zone key (`ALMA_NZ_API_KEY`)
   - Per-campus keys for each Shared Print school (`ALMA_PROD_{campus}`)
   - Sandbox equivalents if needed (`ALMA_SANDBOX_NZ_API_KEY`, `ALMA_SANDBOX_{campus}`)

   The campus code is the part after `01CUNY_` in the Alma Institution Code — for example, `01CUNY_QC` → `ALMA_PROD_QC`.

2. **`data/schools_template.csv`** - Contains school metadata (no API keys):
   - Name, Size (1=largest), Shared Print (Yes/No)
   - Alma Institution Code, OCLC Symbol, MARC Org Code, Primo View ID
   - Chief Librarian Name and Email
   - OCLC Collection ID (7-digit data sync collection ID; required for Phase 5 WorldCat CSV generation)

   A sample file showing the expected format (with placeholder data) is committed as `schools_sample.csv`.

## Web App (Recommended)

The easiest way to use this tool is through the Streamlit web app:

**Option A — Double-click to launch:**
- Open Finder, navigate to the `library-retention` folder
- Double-click **`run_app.command`**
- The app will open in your browser automatically

**Option B — Launch from terminal:**
> **Important:** You must `cd` into the project folder first, or the app won't find its files.

```bash
cd /Users/kristenfredericksen/Library/CloudStorage/OneDrive-CUNY/agentic-projects/library-retention
source venv/bin/activate
streamlit run src/app.py
```

The app has two tabs matching the two-step workflow:
1. **Step 1 — Lookup**: Upload your barcodes Excel file, view results, and download draft emails
2. **Step 2 — Update**: Upload the pending JSON from Step 1, record yes/no/skip for each school, and apply Alma + WorldCat updates

## Command-Line Usage

You can also run the script directly from the terminal.

### Input File Format

Create an Excel file with two columns:

| Barcode | School Code |
|---------|-------------|
| 31699001195116 | 01CUNY_JJ |
| 31228005903455 | 01CUNY_BC |

- **Barcode**: The item's barcode
- **School Code**: The Alma Institution Code of the school leaving retention

### Running the Script

> **Every time you open a new terminal**, `cd` to the project folder and activate the virtual environment:
> ```bash
> cd /Users/kristenfredericksen/Library/CloudStorage/OneDrive-CUNY/agentic-projects/library-retention
> source venv/bin/activate
> ```
> You'll see `(venv)` at the start of your prompt when it's active.

```bash
# Sandbox mode (safe for testing)
python3 src/retention_transfer.py data/your_barcodes.xlsx --sandbox

# Production mode
python3 src/retention_transfer.py data/your_barcodes.xlsx

# Specify a custom output directory for .eml files
python3 src/retention_transfer.py data/your_barcodes.xlsx output/emails --sandbox
```

**Shortcut:** If you set up the `retention` shell alias (see below), just type `retention` instead of the `cd` and `source` commands.

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

3. Generate WorldCat update files (Phase 5):
   - **Taking school CSV** — saved to `output/worldcat/taking/`
     - OCLC Full Format Shared Print file (14 columns)
     - Filename: `<collectionID>.<OCLCsymbol>.sharedprint_retention_transfer_<YYYYMMDD>.csv`
     - Upload via: WorldShare Metadata > My Files > Uploads > file type "Data sync LHR"
     - Only generated for schools with an OCLC Collection ID in the schools CSV
   - **Leaving school instructions** — saved to `output/worldcat/leaving/`
     - Lists all records that need retention commitments removed
     - Includes OCLC numbers, barcodes, and step-by-step removal instructions

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

## Shell Alias (Optional Shortcut)

To avoid typing the full `cd` and `source` commands every time, you can add a shortcut to your terminal. Run this once:

```bash
echo 'alias retention="cd /Users/kristenfredericksen/Library/CloudStorage/OneDrive-CUNY/agentic-projects/library-retention && source venv/bin/activate"' >> ~/.zshrc && source ~/.zshrc
```

After that, just type `retention` in any terminal window to navigate to the project and activate the virtual environment.

## Files Not in Repository

The following files contain sensitive data and are excluded from git:
- `.env` (API keys)
- `data/` folder (school information including librarian contacts, and barcode files)

`schools_sample.csv` (project root) is committed as a format reference with placeholder data.

## License

Internal CUNY use only.
