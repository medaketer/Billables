import streamlit as st

from billables_core import (
    parse_billables_docx,
    load_or_create_workbook,
    add_week_sheet,
    next_week_sheet_name,
    consolidate_workbook,
    workbook_to_bytes,
    WEEK_SHEET_RE,
)

st.set_page_config(page_title="Billables Workbook Tool", page_icon="🧾", layout="centered")

st.title("🧾 Billables Workbook Tool")
st.caption(
    "Turn a weekly Billables Word doc into a tab in your workbook, "
    "and consolidate all the weeks into one summary at month end."
)

tab1, tab2 = st.tabs(["➕ Add a Week", "📊 Consolidate Month"])

# ---------------------------------------------------------------------------
# TAB 1: Add a week
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Add this week's data")
    st.write(
        "Upload the weekly Billables Word document, and (if you have one) "
        "the master workbook you've been building. A new tab will be added "
        "for this week — the rest of the workbook stays untouched."
    )

    docx_file = st.file_uploader("Weekly Billables Word document (.docx)", type=["docx"], key="docx_upload")
    xlsx_file = st.file_uploader(
        "Master workbook (.xlsx) — leave empty to start a brand new one",
        type=["xlsx"],
        key="xlsx_upload_week",
    )

    if docx_file is not None:
        try:
            wb_preview = load_or_create_workbook(xlsx_file)
            suggested_name = next_week_sheet_name(wb_preview)
        except Exception as e:
            st.error(f"Couldn't read that workbook: {e}")
            suggested_name = "Week 1"

        sheet_name = st.text_input("Tab name for this week", value=suggested_name)

        if st.button("Parse & Add Week", type="primary"):
            try:
                entries = parse_billables_docx(docx_file)
                if not entries:
                    st.warning(
                        "No billable entries were found in that document. "
                        "Double-check it follows the usual layout "
                        "(highlighted property name, then date + description + $ amount)."
                    )
                else:
                    xlsx_file.seek(0) if xlsx_file is not None else None
                    wb = load_or_create_workbook(xlsx_file)
                    add_week_sheet(wb, sheet_name.strip(), entries)
                    total = sum(e[3] for e in entries)

                    st.success(f"Added tab '{sheet_name}' with {len(entries)} entries — total ${total:,.2f}")
                    st.dataframe(
                        [{"Property": p, "Date": d.strftime("%m/%d/%y"), "Description": desc[:80] + ("…" if len(desc) > 80 else ""), "Amount": f"${amt:,.2f}"}
                         for p, d, desc, amt in entries],
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.download_button(
                        "⬇️ Download updated workbook",
                        data=workbook_to_bytes(wb),
                        file_name=xlsx_file.name if xlsx_file is not None else "Billables_Workbook.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.info(
                        "Save this file and use it as the 'master workbook' next week "
                        "(and at month end, for consolidating)."
                    )
            except Exception as e:
                st.error(f"Something went wrong parsing that document: {e}")
    else:
        st.info("Upload this week's Word document to get started.")

# ---------------------------------------------------------------------------
# TAB 2: Consolidate
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Consolidate the month")
    st.write(
        "Upload the workbook that has all of this month's 'Week 1', 'Week 2', "
        "etc. tabs in it. This will build (or refresh) a **Consolidated** tab "
        "with every entry, and a **Summary** tab totaling billables by property."
    )

    xlsx_file2 = st.file_uploader("Workbook with Week 1, Week 2, ... tabs (.xlsx)", type=["xlsx"], key="xlsx_upload_month")

    if xlsx_file2 is not None:
        try:
            wb = load_or_create_workbook(xlsx_file2)
            week_sheets = [n for n in wb.sheetnames if WEEK_SHEET_RE.match(n)]
        except Exception as e:
            st.error(f"Couldn't read that workbook: {e}")
            week_sheets = []
            wb = None

        if wb is not None:
            if not week_sheets:
                st.warning("No tabs named 'Week 1', 'Week 2', etc. were found in this workbook.")
            else:
                st.write(f"Found tabs: {', '.join(sorted(week_sheets))}")
                if st.button("Consolidate", type="primary"):
                    try:
                        n_entries, n_props = consolidate_workbook(wb)
                        st.success(
                            f"Consolidated {n_entries} entries across {len(week_sheets)} weeks "
                            f"({n_props} unique properties). Added 'Consolidated' and 'Summary' tabs."
                        )
                        st.download_button(
                            "⬇️ Download consolidated workbook",
                            data=workbook_to_bytes(wb),
                            file_name=xlsx_file2.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"Something went wrong consolidating: {e}")
    else:
        st.info("Upload the month's workbook (with its Week tabs) to consolidate.")

st.divider()
with st.expander("ℹ️ How this works / troubleshooting"):
    st.markdown(
        """
- **Weekly step:** upload the technician's Word doc → app adds a `Week N` tab to your workbook → download and save that file. Next week, upload *that* file back in as the master workbook.
- **Month-end step:** once all the week tabs are in one workbook, use the Consolidate tab to build a combined view and a per-property summary.
- The parser relies on the Word doc's usual formatting: property names are **highlighted**, and each entry line starts with a date (`MM/DD/YY`) and ends with a dollar amount. If the layout of the source document changes, parsing may need adjusting.
- Nothing you upload is stored anywhere — each session only works with the files you provide, and the result is yours to download.
        """
    )
