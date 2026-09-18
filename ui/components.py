"""
ui/components.py

Reusable Streamlit UI components for ALFA.

  - Metric cards bar (total, success, failed, pending)
  - Per-company status cards with badges
  - Interactive preview editor (st.data_editor)
  - ZIP packaging for bulk downloads
"""

import io
import zipfile
import streamlit as st
import pandas as pd
from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# Metric cards
# ---------------------------------------------------------------------------

def render_metrics_bar(total: int, success: int, failed: int, pending: int):
    """Render a row of four styled metric cards."""
    cols = st.columns(4)
    cards = [
        ("total", ":material/description: Total", total),
        ("success", ":material/check_circle: Succeeded", success),
        ("failed", ":material/error: Failed", failed),
        ("pending", ":material/schedule: Pending", pending),
    ]
    for col, (css_class, label, value) in zip(cols, cards):
        with col:
            st.markdown(
                f"""
                <div class="metric-card {css_class}">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------------------
# Status badges
# ---------------------------------------------------------------------------

_STATUS_LABELS = {
    "pending": "Pending",
    "processing": "Processing",
    "success": "Success",
    "failed": "Failed",
}


def render_company_card(name: str, status: str, mode: str = "", error: str = ""):
    """Render a single company status card."""
    badge_label = _STATUS_LABELS.get(status, status)
    mode_text = f'<span class="company-mode">({mode})</span>' if mode else ""
    error_html = f'<div style="color:#f87171;font-size:0.8rem;margin-top:0.4rem;">{error}</div>' if error else ""

    st.markdown(
        f"""
        <div class="company-card">
            <div>
                <span class="company-name">{name}</span>
                {mode_text}
                {error_html}
            </div>
            <span class="status-badge {status}">{badge_label}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Preview editor
# ---------------------------------------------------------------------------

def render_preview_editor(data: dict, company_name: str) -> dict:
    """
    Show an interactive editor for the extracted data so the user can
    review and override values before Excel generation.

    Returns the (potentially modified) data dict.
    """
    f_ = data["fields"]

    st.subheader(f"Review: {company_name}", anchor=False)

    with st.expander("Financial summary", icon=":material/bar_chart:", expanded=False):
        fin_df = pd.DataFrame([
            {"Field": "Standalone Turnover FY25 (₹L)", "Value": f_["standalone_turnover_fy25_lakhs"]["value"],
             "Source": f_["standalone_turnover_fy25_lakhs"]["source"]},
            {"Field": "Standalone Turnover FY24 (₹L)", "Value": f_["standalone_turnover_fy24_lakhs"]["value"],
             "Source": f_["standalone_turnover_fy24_lakhs"]["source"]},
            {"Field": "Total Cost FY25 (₹L)", "Value": f_["standalone_total_cost_fy25_lakhs"]["value"],
             "Source": f_["standalone_total_cost_fy25_lakhs"]["source"]},
            {"Field": "Total Cost FY24 (₹L)", "Value": f_["standalone_total_cost_fy24_lakhs"]["value"],
             "Source": f_["standalone_total_cost_fy24_lakhs"]["source"]},
            {"Field": "Cash FY25 (₹L)", "Value": f_["cash_fy25_lakhs"]["value"],
             "Source": f_["cash_fy25_lakhs"]["source"]},
            {"Field": "Cash FY24 (₹L)", "Value": f_["cash_fy24_lakhs"]["value"],
             "Source": f_["cash_fy24_lakhs"]["source"]},
        ])
        edited_fin = st.data_editor(fin_df, key=f"fin_{company_name}", num_rows="fixed")

        # Write back edits
        fin_fields = [
            "standalone_turnover_fy25_lakhs", "standalone_turnover_fy24_lakhs",
            "standalone_total_cost_fy25_lakhs", "standalone_total_cost_fy24_lakhs",
            "cash_fy25_lakhs", "cash_fy24_lakhs",
        ]
        for i, field_name in enumerate(fin_fields):
            f_[field_name]["value"] = edited_fin.iloc[i]["Value"]

    with st.expander("Shareholding", icon=":material/group:", expanded=False):
        sh_rows = f_["shareholding"]["rows"]
        sh_df = pd.DataFrame(sh_rows)
        edited_sh = st.data_editor(sh_df, key=f"sh_{company_name}", num_rows="dynamic")
        f_["shareholding"]["rows"] = edited_sh.to_dict("records")

    with st.expander("Related party transactions", icon=":material/handshake:", expanded=False):
        rpt_items = f_["related_party_transactions_lakhs"]["items"]
        if rpt_items:
            rpt_df = pd.DataFrame(rpt_items)
            edited_rpt = st.data_editor(rpt_df, key=f"rpt_{company_name}", num_rows="dynamic")
            f_["related_party_transactions_lakhs"]["items"] = edited_rpt.to_dict("records")
        else:
            st.info("No related party transactions extracted.", icon=":material/info:")

    with st.expander("Litigation", icon=":material/gavel:", expanded=False):
        lit_items = f_["litigation"]["items"]
        if lit_items:
            lit_df = pd.DataFrame(lit_items)
            edited_lit = st.data_editor(lit_df, key=f"lit_{company_name}", num_rows="dynamic")
            f_["litigation"]["items"] = edited_lit.to_dict("records")
        else:
            st.info("No litigation items extracted.", icon=":material/info:")

    return data


# ---------------------------------------------------------------------------
# ZIP packaging
# ---------------------------------------------------------------------------

def create_zip_download(file_dict: Dict[str, bytes]) -> bytes:
    """
    Package multiple Excel files into a single ZIP archive.

    Args:
        file_dict: mapping of filename → file bytes

    Returns:
        ZIP file bytes ready for st.download_button
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in file_dict.items():
            zf.writestr(name, data)
    return buffer.getvalue()


def render_download_section(results: Dict[str, bytes], errors: Dict[str, str]):
    """
    Render individual download buttons for each company, plus a bulk ZIP
    download button if there are multiple successful results.
    """
    if not results:
        st.warning("No successfully processed files to download.")
        return

    st.subheader("Downloads", anchor=False)

    # Individual buttons
    for i, (name, data) in enumerate(results.items()):
        st.download_button(
            label=f"Download: {name}",
            data=data,
            file_name=f"{name}_Populated_BD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_btn_{i}",
            icon=":material/download:",
        )

    # Bulk ZIP
    if len(results) > 1:
        zip_files = {f"{name}_Populated_BD.xlsx": data for name, data in results.items()}
        zip_bytes = create_zip_download(zip_files)
        st.download_button(
            label=f"Download all ({len(results)} files) as ZIP",
            data=zip_bytes,
            file_name="ALFA_Populated_BD_All.zip",
            mime="application/zip",
            key="dl_zip",
            type="primary",
            icon=":material/folder_zip:",
        )
