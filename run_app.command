#!/bin/bash
# ──────────────────────────────────────────────────────────────────
# CUNY Shared Print — Retention Transfer App
# Double-click this file in Finder to open the web app.
# ──────────────────────────────────────────────────────────────────

# Move to the project folder (same folder as this script)
cd "$(dirname "$0")"

# Activate the virtual environment
source venv/bin/activate

# Open the app in the browser (Streamlit does this automatically)
streamlit run src/app.py

# Keep the terminal window open if something goes wrong
read -p "Press Enter to close this window…"
