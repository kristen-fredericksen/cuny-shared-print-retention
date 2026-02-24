"""
CUNY Shared Print Monograph Trust - Retention Transfer  (Streamlit UI)

Run with:
    cd /Users/kristenfredericksen/Library/CloudStorage/OneDrive-CUNY/agentic-projects/library-retention
    source venv/bin/activate
    streamlit run src/app.py
"""

import os
import sys
import json
import tempfile
import zipfile
import io
from datetime import datetime
from pathlib import Path

import streamlit as st

# ---------------------------------------------------------------------------
# Make sure the project root is on the path so we can import retention_transfer
# ---------------------------------------------------------------------------
_HERE = Path(__file__).parent
_ROOT = _HERE.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# Redirect sys.exit to a Streamlit-friendly exception so the app never crashes
import builtins as _builtins

class _AppExit(Exception):
    pass

_real_exit = sys.exit
def _patched_exit(code=0):
    raise _AppExit(f"sys.exit({code})")

sys.exit = _patched_exit

try:
    from src.retention_transfer import (
        get_config,
        load_schools,
        read_barcodes,
        process_barcodes,
        print_summary,
        print_draft_emails,
        save_pending_transfers,
        load_pending_transfers,
        _handle_decline,
        _re_verify_leaving_school_ids,
        _re_verify_taking_school_ids,
        update_leaving_school_item,
        update_leaving_school_holdings,
        update_taking_school_item,
        update_taking_school_holdings,
        generate_worldcat_taking_csv,
        generate_worldcat_leaving_instructions,
    )
    _IMPORT_OK = True
    _IMPORT_ERROR = None
except Exception as _e:
    _IMPORT_OK = False
    _IMPORT_ERROR = str(_e)
finally:
    sys.exit = _real_exit  # restore

import pandas as pd

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="CUNY Shared Print — Retention Transfer",
    page_icon="📚",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session-state helpers
# ---------------------------------------------------------------------------
def _ss(key, default=None):
    """Get a session-state value, initialising it to default if absent."""
    if key not in st.session_state:
        st.session_state[key] = default
    return st.session_state[key]


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _zip_directory(directory: str) -> bytes:
    """Zip all files in `directory` and return raw bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(directory):
            for fname in files:
                full_path = os.path.join(root, fname)
                arcname = os.path.relpath(full_path, directory)
                zf.write(full_path, arcname)
    return buf.getvalue()


def _read_file_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def _status_badge(status: str) -> str:
    """Return a coloured emoji + label for a status string."""
    return {
        "replacement_found": "✅ Replacement found",
        "awaiting_reply":    "⏳ Awaiting reply",
        "completed":         "✅ Completed",
        "no_replacement":    "⚠️ No replacement",
        "not_found":         "❌ Not found",
        "ineligible":        "🚫 Ineligible",
        "error":             "🔴 Error",
    }.get(status, f"❓ {status}")


def _school_name(code: str | None, schools: dict) -> str:
    if not code:
        return "—"
    return schools.get(code, {}).get("name", code)


# ---------------------------------------------------------------------------
# Capture stdout so we can display it inside the app
# ---------------------------------------------------------------------------
import contextlib
import io as _io

@contextlib.contextmanager
def _capture_output():
    """Context manager that captures stdout into a string."""
    buf = _io.StringIO()
    with contextlib.redirect_stdout(buf):
        yield buf


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

def main():
    # ── Header ──────────────────────────────────────────────────────────────
    st.title("📚 CUNY Shared Print — Retention Transfer")
    st.caption("Transfer retention commitments when a school can no longer retain a book.")

    if not _IMPORT_OK:
        st.error(f"**Could not import retention_transfer.py.**\n\n```\n{_IMPORT_ERROR}\n```")
        st.info("Make sure you launched the app from the project root with the venv active.")
        return

    # ── Sidebar: environment ─────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Settings")
        sandbox = st.toggle("Sandbox mode", value=True,
                            help="Use sandbox API keys and sandbox schools file. "
                                 "Turn off only when ready for production.")
        if sandbox:
            st.warning("🧪 Sandbox mode — changes go to the sandbox only.")
        else:
            st.error("⚠️ Production mode — changes affect live Alma records.")

        st.divider()

        # Load config + schools (cached in session state per sandbox toggle)
        cache_key = f"config_schools_{sandbox}"
        if cache_key not in st.session_state:
            try:
                with _capture_output() as out:
                    config  = get_config(sandbox=sandbox)
                    schools = load_schools(config["schools_file"])
                st.session_state[cache_key] = (config, schools)
                _env_log = out.getvalue()
            except _AppExit as e:
                st.error(f"Configuration error: {e}")
                st.info("Check your `.env` file — it needs the correct API keys.")
                st.stop()

        config, schools = st.session_state[cache_key]

        st.success(f"✓ Loaded {len(schools)} schools")
        st.caption(f"API: `{config['base_url']}`")

    # ── Tabs ────────────────────────────────────────────────────────────────
    tab1, tab2 = st.tabs(["📋 Step 1 — Lookup", "✏️ Step 2 — Update"])

    # =========================================================================
    # TAB 1 — LOOKUP
    # =========================================================================
    with tab1:
        st.header("Step 1 — Look up barcodes & select replacement schools")
        st.markdown(
            "Upload an Excel file with the barcodes of items the leaving school "
            "can no longer retain.  The script will look up each item in Alma, "
            "select a replacement school, and generate draft emails.  "
            "**No records are changed in this step.**"
        )

        st.subheader("Required Excel format")
        st.dataframe(
            pd.DataFrame({
                "Barcode":     ["39016013760757", "39016011424125"],
                "School Code": ["01CUNY_BC", "01CUNY_CC"],
            }),
            use_container_width=False,
            hide_index=True,
        )

        # ── File upload ──────────────────────────────────────────────────────
        uploaded_xlsx = st.file_uploader(
            "Upload barcodes Excel file (.xlsx)",
            type=["xlsx", "xls"],
            key="barcode_upload",
        )

        # ── Output directory ─────────────────────────────────────────────────
        output_dir_default = str(_ROOT / "output" / datetime.now().strftime("run_%Y%m%d_%H%M%S"))
        output_dir_input = st.text_input(
            "Output folder (where emails and pending JSON will be saved)",
            value=output_dir_default,
            help="A new folder will be created automatically. You can change this path.",
        )

        # ── Run button ───────────────────────────────────────────────────────
        run_lookup = st.button("🔍 Run Lookup", type="primary", disabled=(uploaded_xlsx is None))

        if run_lookup and uploaded_xlsx is not None:
            # Save uploaded file to a temp location so retention_transfer can read it
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_xlsx.getbuffer())
                tmp_path = tmp.name

            output_dir = output_dir_input.strip() or output_dir_default
            email_dir  = os.path.join(output_dir, "emails")

            log_lines    = []
            results      = None
            pending_path = None
            lookup_error = None

            # Progress bar lives outside the spinner so it can update in real time
            progress_bar  = st.progress(0, text="Reading barcodes…")
            status_text   = st.empty()

            def _update_progress(current, total, barcode, status):
                pct = current / total
                progress_bar.progress(pct, text=f"[{current}/{total}]")
                badge = _status_badge(status)
                status_text.caption(f"Last: {barcode} → {badge}")

            try:
                with _capture_output() as out:
                    items = read_barcodes(tmp_path)

                if not items:
                    lookup_error = "No items found in the uploaded file. Check the column names."
                else:
                    progress_bar.progress(0, text=f"Looking up {len(items)} items in Alma…")

                    with _capture_output() as out:
                        results = process_barcodes(items, schools, config,
                                                   progress_callback=_update_progress)
                    log_lines.append(out.getvalue())

                    progress_bar.progress(1.0, text="Generating emails…")
                    with _capture_output() as out:
                        print_summary(results, schools)
                    log_lines.append(out.getvalue())

                    with _capture_output() as out:
                        print_draft_emails(results, schools, email_dir)
                    log_lines.append(out.getvalue())

                    pending_path = save_pending_transfers(results, output_dir, schools)
                    log_lines.append(f"\n✓ Pending file saved: {pending_path}\n")

                    progress_bar.progress(1.0, text="Done!")

            except _AppExit as e:
                lookup_error = str(e)
            except Exception as e:
                lookup_error = f"Unexpected error: {e}"
            finally:
                os.unlink(tmp_path)

            if lookup_error:
                st.error(lookup_error)
            else:
                # ── Summary banner ───────────────────────────────────────────
                found     = [r for r in results if r["status"] == "replacement_found"]
                no_repl   = [r for r in results if r["status"] == "no_replacement"]
                not_found = [r for r in results if r["status"] == "not_found"]
                ineligible = [r for r in results if r["status"] == "ineligible"]
                errors    = [r for r in results if r["status"] == "error"]

                st.success(f"✓ Lookup complete — {len(results)} item(s) processed, "
                           f"{len(found)} replacement(s) found.")

                # ── Replacements found ───────────────────────────────────────
                if found:
                    st.divider()
                    st.subheader(f"✅ Replacement found ({len(found)})")
                    st.dataframe(
                        pd.DataFrame([{
                            "Barcode":     r.get("barcode", ""),
                            "Title":       (r.get("title") or "")[:60],
                            "Leaving":     _school_name(r.get("leaving_school"), schools),
                            "Replacement": _school_name(r.get("replacement_school"), schools),
                            "All holders": ", ".join(
                                _school_name(c, schools)
                                for c in r.get("holding_institutions", [])
                            ),
                        } for r in found]),
                        use_container_width=True, hide_index=True,
                    )

                # ── Items skipped (wrong status) ──────────────────────────────
                if ineligible:
                    st.divider()
                    st.subheader(f"🚫 Skipped — not in 'Item in place' status ({len(ineligible)})")
                    st.dataframe(
                        pd.DataFrame([{
                            "Barcode": r.get("barcode", ""),
                            "Title":   (r.get("title") or "")[:60],
                            "Status":  r.get("item_status", "unknown status"),
                        } for r in ineligible]),
                        use_container_width=True, hide_index=True,
                    )

                # ── Barcodes not found ────────────────────────────────────────
                if not_found:
                    st.divider()
                    st.subheader(f"❌ Not found in Alma ({len(not_found)})")
                    st.dataframe(
                        pd.DataFrame([{
                            "Barcode": r.get("barcode", ""),
                            "Leaving": _school_name(r.get("leaving_school"), schools),
                            "Details": r.get("error", ""),
                        } for r in not_found]),
                        use_container_width=True, hide_index=True,
                    )

                # ── Errors ────────────────────────────────────────────────────
                if errors:
                    st.divider()
                    st.subheader(f"🔴 Errors ({len(errors)})")
                    st.dataframe(
                        pd.DataFrame([{
                            "Barcode": r.get("barcode", ""),
                            "Leaving": _school_name(r.get("leaving_school"), schools),
                            "Error":   r.get("error", ""),
                        } for r in errors]),
                        use_container_width=True, hide_index=True,
                    )

                # ── No eligible replacement ───────────────────────────────────
                if no_repl:
                    st.divider()
                    st.subheader(f"⚠️ No eligible replacement — flagged for withdrawal review ({len(no_repl)})")
                    st.dataframe(
                        pd.DataFrame([{
                            "Barcode": r.get("barcode", ""),
                            "Title":   (r.get("title") or "Unknown")[:60],
                            "Reason":  r.get("no_replacement_reason", ""),
                        } for r in no_repl]),
                        use_container_width=True, hide_index=True,
                    )

                # ── Full results table ────────────────────────────────────────
                st.divider()
                st.subheader(f"📋 All results ({len(results)})")
                rows = []
                for r in results:
                    detail = (
                        r.get("error")
                        or r.get("item_status")
                        or r.get("no_replacement_reason")
                        or ""
                    )
                    if not detail:
                        holders = r.get("holding_institutions", [])
                        if holders:
                            holder_names = [_school_name(c, schools) for c in holders]
                            detail = f"Held by: {', '.join(holder_names)}"
                    rows.append({
                        "Barcode":     r.get("barcode", ""),
                        "Title":       (r.get("title") or "")[:60],
                        "Status":      _status_badge(r.get("status", "")),
                        "Leaving":     _school_name(r.get("leaving_school"), schools),
                        "Replacement": _school_name(r.get("replacement_school"), schools),
                        "Details":     detail,
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                # ── Downloads ────────────────────────────────────────────────
                st.subheader("📥 Download outputs")
                col1, col2 = st.columns(2)

                # Pending JSON
                with col1:
                    if pending_path and os.path.exists(pending_path):
                        with open(pending_path, "rb") as f:
                            st.download_button(
                                "⬇ Download pending JSON",
                                data=f.read(),
                                file_name=os.path.basename(pending_path),
                                mime="application/json",
                            )
                        st.caption(f"Save this file — you'll need it for Step 2.")

                # Draft emails zip
                with col2:
                    if os.path.isdir(email_dir) and any(
                        f.endswith(".eml") for f in os.listdir(email_dir)
                    ):
                        zip_bytes = _zip_directory(email_dir)
                        st.download_button(
                            "⬇ Download draft emails (.zip)",
                            data=zip_bytes,
                            file_name="draft_emails.zip",
                            mime="application/zip",
                        )
                        st.caption("Open each .eml file in your email client, review, and send.")

                # ── Log ──────────────────────────────────────────────────────
                with st.expander("📜 Full log", expanded=False):
                    st.code("".join(log_lines), language=None)

                # ── Next steps ───────────────────────────────────────────────
                st.info(
                    "**Next steps:**\n"
                    "1. Download and send the draft emails.\n"
                    "2. Wait for replies from the chief librarians.\n"
                    "3. When you have replies, come back and use **Step 2 — Update**."
                )

    # =========================================================================
    # TAB 2 — UPDATE
    # =========================================================================
    with tab2:
        st.header("Step 2 — Record replies & update Alma / WorldCat")
        st.markdown(
            "Upload the pending JSON file from Step 1.  For each item that is "
            "awaiting a reply, indicate whether the school agreed, declined, or "
            "if you want to skip it for now."
        )

        # ── Load pending JSON ────────────────────────────────────────────────
        uploaded_json = st.file_uploader(
            "Upload pending JSON file",
            type=["json"],
            key="pending_upload",
        )

        # Let the user also point to a file already on disk
        st.caption("— or —")
        json_path_input = st.text_input(
            "Path to pending JSON file on this computer",
            placeholder="/path/to/output/pending_20260220_143012.json",
            key="json_path_input",
        )

        # Resolve the pending file source
        pending_payload = None
        pending_source  = None   # "upload" | "path"
        pending_json_path = None

        if uploaded_json is not None:
            try:
                pending_payload = json.load(uploaded_json)
                # Save to temp file so we can overwrite it later
                with tempfile.NamedTemporaryFile(
                    suffix=".json", delete=False, mode="w", encoding="utf-8"
                ) as tmp:
                    json.dump(pending_payload, tmp, indent=2, default=str)
                    pending_json_path = tmp.name
                pending_source = "upload"
            except Exception as e:
                st.error(f"Could not read JSON file: {e}")

        elif json_path_input.strip():
            p = json_path_input.strip()
            if os.path.exists(p):
                try:
                    with open(p, encoding="utf-8") as f:
                        pending_payload = json.load(f)
                    pending_json_path = p
                    pending_source = "path"
                except Exception as e:
                    st.error(f"Could not read file: {e}")
            else:
                st.error(f"File not found: {p}")

        if pending_payload is not None:
            if pending_payload.get("version") != 1 or "items" not in pending_payload:
                st.error("This doesn't look like a valid pending-transfers file.")
                pending_payload = None

        # ── Show items ───────────────────────────────────────────────────────
        if pending_payload is not None:
            items = pending_payload["items"]
            created = pending_payload.get("created", "unknown")[:19].replace("T", " ")

            st.success(f"✓ Loaded {len(items)} item(s)  (lookup date: {created})")

            # Quick status summary
            tally = {}
            for it in items:
                tally[it["status"]] = tally.get(it["status"], 0) + 1

            cols = st.columns(len(tally) or 1)
            for col, (status, count) in zip(cols, tally.items()):
                col.metric(_status_badge(status), count)

            st.divider()

            # ── Awaiting-reply items ─────────────────────────────────────────
            awaiting = [it for it in items if it["status"] == "awaiting_reply"]

            if not awaiting:
                st.info("No items are currently awaiting a reply.  Nothing to do right now.")
            else:
                st.subheader(f"⏳ Items awaiting a reply ({len(awaiting)})")
                st.markdown(
                    "For each item below, choose whether the proposed school "
                    "**agreed**, **declined**, or you want to **skip** for now."
                )

                # Output dir for this update run
                if pending_source == "path":
                    update_output_default = str(Path(pending_json_path).parent)
                else:
                    update_output_default = str(_ROOT / "output")

                update_output = st.text_input(
                    "Output folder for WorldCat files and updated emails",
                    value=update_output_default,
                    key="update_output_dir",
                )

                # ── Per-item decision widgets ────────────────────────────────
                # We store decisions in session state so they survive reruns
                if "item_decisions" not in st.session_state:
                    st.session_state["item_decisions"] = {}

                decisions = st.session_state["item_decisions"]

                for it in awaiting:
                    barcode      = it.get("barcode", "?")
                    title        = it.get("title") or "Unknown title"
                    leaving_code = it.get("leaving_school", "")
                    taking_code  = it.get("proposed_school", "")
                    declined     = it.get("declined_schools", [])

                    with st.container(border=True):
                        c1, c2 = st.columns([3, 1])
                        with c1:
                            st.markdown(f"**{title[:80]}**")
                            st.caption(
                                f"Barcode: `{barcode}`  |  "
                                f"Leaving: {_school_name(leaving_code, schools)}  |  "
                                f"Proposed: **{_school_name(taking_code, schools)}**"
                            )
                            if declined:
                                declined_names = ", ".join(
                                    _school_name(s, schools) for s in declined
                                )
                                st.caption(f"Previously declined: {declined_names}")
                        with c2:
                            key = f"decision_{barcode}"
                            decisions[barcode] = st.radio(
                                "Reply",
                                options=["(not yet decided)", "✅ Yes — agreed", "❌ No — declined", "⏭ Skip for now"],
                                key=key,
                                label_visibility="collapsed",
                            )

                # ── Run Updates button ───────────────────────────────────────
                st.divider()

                ready_count = sum(
                    1 for v in decisions.values()
                    if v in ("✅ Yes — agreed", "❌ No — declined")
                )

                run_update = st.button(
                    f"▶️ Apply decisions ({ready_count} item(s) ready)",
                    type="primary",
                    disabled=(ready_count == 0),
                )

                if run_update:
                    output_dir = update_output.strip() or update_output_default
                    email_dir  = os.path.join(output_dir, "emails")
                    os.makedirs(email_dir, exist_ok=True)

                    update_log  = []
                    n_completed = 0
                    n_declined  = 0
                    n_skipped   = 0
                    n_errors    = 0
                    new_emails  = []

                    progress = st.progress(0, text="Applying decisions…")
                    status_area = st.empty()

                    for idx, it in enumerate(awaiting):
                        barcode  = it.get("barcode", "?")
                        decision = decisions.get(barcode, "(not yet decided)")
                        total    = len(awaiting)

                        progress.progress((idx) / total, text=f"Processing {barcode}…")

                        if decision == "(not yet decided)" or decision == "⏭ Skip for now":
                            n_skipped += 1
                            update_log.append(f"[{barcode}] Skipped.")
                            continue

                        if decision == "❌ No — declined":
                            msg = _handle_decline(it, schools, email_dir)
                            n_declined += 1
                            update_log.append(f"[{barcode}] Declined → {msg}")
                            # Track new email if one was created
                            if "New email:" in msg:
                                new_emails.append(msg.split("New email:")[-1].strip())
                            continue

                        # ── Yes — agreed ─────────────────────────────────────
                        taking_code  = it.get("proposed_school", "")
                        leaving_code = it.get("leaving_school", "")

                        update_log.append(f"[{barcode}] {_school_name(taking_code, schools)} agreed. Re-verifying…")

                        # Re-verify leaving school IDs
                        l_mms_id, l_holding_id, l_item_pid, warn = _re_verify_leaving_school_ids(
                            it, schools, config
                        )
                        if l_mms_id is None:
                            n_errors += 1
                            update_log.append(f"  ✗ Could not verify leaving school IDs: {warn}")
                            continue
                        if warn:
                            update_log.append(f"  ⚠ IDs changed: {warn}")

                        # Re-verify taking school IDs
                        t_iz_mms_id, t_holding_id, t_item_pid, err = _re_verify_taking_school_ids(
                            it, taking_code, schools, config
                        )
                        if err:
                            n_errors += 1
                            update_log.append(f"  ✗ Could not verify taking school IDs: {err}")
                            continue

                        # Phase 3: Update leaving school
                        leaving_school = schools.get(leaving_code, {})
                        l_api_key      = leaving_school.get("api_key", "")

                        item_ok, item_msg = update_leaving_school_item(
                            l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
                        )
                        if item_ok:
                            update_log.append(f"  ✓ Leaving item: {item_msg}")
                        else:
                            n_errors += 1
                            update_log.append(f"  ✗ Leaving item failed: {item_msg}")
                            continue

                        _, holdings_msg = update_leaving_school_holdings(
                            l_mms_id, l_holding_id, l_item_pid, l_api_key, config["base_url"]
                        )
                        update_log.append(f"  ✓ Leaving holdings: {holdings_msg}")

                        # Phase 4: Update taking school
                        taking_school = schools.get(taking_code, {})
                        t_api_key     = taking_school.get("api_key", "")
                        marc_org_code = taking_school.get("marc_org_code", "") or ""

                        item_ok, item_msg = update_taking_school_item(
                            t_iz_mms_id, t_holding_id, t_item_pid, t_api_key, config["base_url"]
                        )
                        if item_ok:
                            update_log.append(f"  ✓ Taking item: {item_msg}")
                        else:
                            n_errors += 1
                            update_log.append(f"  ✗ Taking item failed: {item_msg}")
                            continue

                        _, holdings_msg = update_taking_school_holdings(
                            t_iz_mms_id, t_holding_id, marc_org_code, t_api_key, config["base_url"]
                        )
                        update_log.append(f"  ✓ Taking holdings: {holdings_msg}")

                        # Phase 5: WorldCat CSV
                        worldcat_result = {
                            "status":             "replacement_found",
                            "barcode":            barcode,
                            "title":              it.get("title"),
                            "replacement_school": taking_code,
                            "leaving_school":     leaving_code,
                            "bib_info":           it.get("bib_info"),
                        }
                        wc_files, wc_skipped = generate_worldcat_taking_csv(
                            [worldcat_result], schools, output_dir
                        )
                        if wc_files:
                            update_log.append(f"  ✓ WorldCat CSV: {wc_files[0]}")
                        elif wc_skipped:
                            update_log.append(f"  ⚠ WorldCat CSV skipped: {wc_skipped[0]['reason']}")

                        wc_instr = generate_worldcat_leaving_instructions(
                            [worldcat_result], schools, output_dir
                        )
                        if wc_instr:
                            update_log.append(f"  ✓ WorldCat leaving instructions: {wc_instr}")

                        # Mark completed in the payload
                        it["status"]         = "completed"
                        it["completed_date"] = datetime.now().isoformat()
                        it["taking_school"]  = taking_code
                        n_completed += 1
                        update_log.append(f"  ✅ Transfer complete.")

                    progress.progress(1.0, text="Done!")

                    # Save updated JSON
                    if pending_source == "path":
                        save_path = pending_json_path
                    else:
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        save_path = os.path.join(output_dir, f"pending_{ts}.json")

                    os.makedirs(os.path.dirname(save_path) if os.path.dirname(save_path) else ".", exist_ok=True)
                    with open(save_path, "w", encoding="utf-8") as f:
                        json.dump(pending_payload, f, indent=2, default=str)
                    update_log.append(f"\n✓ Pending file updated: {save_path}")

                    # ── Results banner ───────────────────────────────────────
                    st.success(
                        f"Done!  ✅ {n_completed} completed  |  "
                        f"❌ {n_declined} declined (new emails sent)  |  "
                        f"⏭ {n_skipped} skipped  |  "
                        f"🔴 {n_errors} errors"
                    )

                    # ── Download outputs ─────────────────────────────────────
                    st.subheader("📥 Download outputs")
                    dl_cols = st.columns(3)

                    # Updated pending JSON
                    with dl_cols[0]:
                        if os.path.exists(save_path):
                            with open(save_path, "rb") as f:
                                st.download_button(
                                    "⬇ Updated pending JSON",
                                    data=f.read(),
                                    file_name=os.path.basename(save_path),
                                    mime="application/json",
                                )

                    # New draft emails (for declined schools)
                    with dl_cols[1]:
                        if os.path.isdir(email_dir) and any(
                            f.endswith(".eml") for f in os.listdir(email_dir)
                        ):
                            zip_bytes = _zip_directory(email_dir)
                            st.download_button(
                                "⬇ New draft emails (.zip)",
                                data=zip_bytes,
                                file_name="draft_emails.zip",
                                mime="application/zip",
                            )

                    # WorldCat files
                    with dl_cols[2]:
                        if os.path.isdir(output_dir):
                            wc_files = [
                                f for f in os.listdir(output_dir)
                                if f.endswith(".csv") or f.endswith(".txt")
                            ]
                            if wc_files:
                                wc_zip = _zip_directory(output_dir)
                                st.download_button(
                                    "⬇ WorldCat files (.zip)",
                                    data=wc_zip,
                                    file_name="worldcat_files.zip",
                                    mime="application/zip",
                                )

                    # ── Log ──────────────────────────────────────────────────
                    with st.expander("📜 Full log", expanded=True):
                        st.code("\n".join(update_log), language=None)

                    # ── Next steps ───────────────────────────────────────────
                    still_awaiting = [it for it in items if it["status"] == "awaiting_reply"]
                    if still_awaiting:
                        st.info(
                            f"**{len(still_awaiting)} item(s) still awaiting a reply.**\n\n"
                            "Download the updated pending JSON above, then re-upload it "
                            "when you have more replies."
                        )
                    elif all(it["status"] == "completed" for it in items):
                        st.balloons()
                        st.success("🎉 All items have been transferred!")

        # ── Completed / no-replacement summary (always shown if payload loaded) ──
        if pending_payload is not None:
            completed_items = [it for it in items if it["status"] == "completed"]
            no_replace_items = [it for it in items if it["status"] == "no_replacement"]

            if completed_items or no_replace_items:
                st.divider()
                if completed_items:
                    with st.expander(f"✅ Completed transfers ({len(completed_items)})", expanded=False):
                        rows = []
                        for it in completed_items:
                            rows.append({
                                "Barcode":       it.get("barcode", ""),
                                "Title":         (it.get("title") or "")[:60],
                                "From":          _school_name(it.get("leaving_school"), schools),
                                "To":            _school_name(it.get("taking_school"), schools),
                                "Completed":     (it.get("completed_date") or "")[:10],
                            })
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                if no_replace_items:
                    with st.expander(
                        f"⚠️ Items needing withdrawal review ({len(no_replace_items)})",
                        expanded=True,
                    ):
                        st.warning(
                            "All eligible replacement schools have declined these items.  "
                            "They need manual review to decide whether to withdraw them from "
                            "the Shared Print program."
                        )
                        rows = []
                        for it in no_replace_items:
                            rows.append({
                                "Barcode":  it.get("barcode", ""),
                                "Title":    (it.get("title") or "")[:60],
                                "Leaving":  _school_name(it.get("leaving_school"), schools),
                                "Declined": ", ".join(
                                    _school_name(s, schools)
                                    for s in it.get("declined_schools", [])
                                ),
                            })
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
