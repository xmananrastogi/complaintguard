"""
streamlit dashboard for the complaint sla analyser.
upload an excel, see results, filter by vendor, export data.

run with:
    streamlit run app/main.py
"""

import os
import sys
import tempfile
from io import BytesIO

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# make sure we can import from the app package
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app.processor import process_excel
from app.database import init_db, save_upload, get_upload_history

# --------------------------------------------------------------
# page config
# --------------------------------------------------------------
st.set_page_config(
    page_title="ComplaintGuard",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------
# colour scheme — matches what the vba macro used
# --------------------------------------------------------------
COLORS = {
    "Early": "#c6efce",
    "Delayed": "#ffc7ce",
    "On Time": "#c6e0ff",
    "Pending": "#ffeb9c",
}


def style_status(val: str) -> str:
    bg = COLORS.get(val, "white")
    return f"background-color: {bg}; font-weight: bold"


def run():
    init_db()
    st.title("🛡️ ComplaintGuard")
    st.markdown("SLA compliance analyser for IOCL vendor complaints.")

    # ----------------------------------------------------------
    # sidebar — file upload + history
    # ----------------------------------------------------------
    with st.sidebar:
        st.header("Upload")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=["xlsx", "xls", "xlsm"],
            help="The IOCL complaint export with Complaint ID, SLA, Complaint DateTime, Vendor Close DateTime columns.",
        )

        st.divider()
        st.header("History")
        history = get_upload_history()
        if history:
            for h in history[:5]:
                st.caption(f"{h['filename']} — {h['uploaded_at'][:10]}")
                st.caption(f"  {h['total_rows']} rows | ₹{h['total_penalty']:,.0f}")
        else:
            st.caption("No previous uploads yet.")

    # ----------------------------------------------------------
    # main area
    # ----------------------------------------------------------
    if uploaded_file is None:
        st.info("Upload an Excel file to get started.")
        st.markdown("""
        **Expected columns:**
        - `Complaint ID` — unique identifier
        - `Complaint Resolution Time` — SLA as text ("48 hours", "24 hours")
        - `Complaint DateTime` — when the complaint was logged
        - `Vendor Close DateTime` — when the vendor closed it
        - `Vendor Code` — vendor name
        - `RO Code` / `RO Name` — dealer info
        - `Vendor Remarks` — used to detect auto-closed complaints
        """)

        # show sample screenshot area
        st.markdown("---")
        st.markdown("**Need a sample?** The original file is included in `data/sample/`.")
        return

    # ----------------------------------------------------------
    # process the uploaded file
    # ----------------------------------------------------------
    with st.spinner("Processing complaints..."):
        # save to temp file so pandas can read it
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            records, summary = process_excel(tmp_path)
        except Exception as e:
            st.error(f"Failed to process file: {e}")
            st.info("Make sure the file has the expected columns. Check the column names in `processor.py`.")
            os.unlink(tmp_path)
            return

        os.unlink(tmp_path)

    if not records:
        st.warning("No valid complaint records found in the file.")
        return

    # save to database
    upload_id = save_upload(uploaded_file.name, records, summary)
    df = pd.DataFrame(records)

    # ----------------------------------------------------------
    # summary cards
    # ----------------------------------------------------------
    st.subheader("Summary")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Complaints", summary["total"])
    col2.metric("Early", summary["early"], delta=None)
    col3.metric("Delayed", summary["delayed"], delta=None)
    col4.metric("Pending", summary["pending"], delta=None)
    col5.metric("Total Penalty", f"₹{summary['total_penalty']:,.0f}")

    # ----------------------------------------------------------
    # charts
    # ----------------------------------------------------------
    st.subheader("Charts")

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        # status distribution pie
        status_counts = df["status"].value_counts().reset_index()
        status_counts.columns = ["status", "count"]
        fig = px.pie(
            status_counts,
            values="count",
            names="status",
            title="Complaint Status Breakdown",
            color="status",
            color_discrete_map=COLORS,
        )
        fig.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        # vendor penalty bar
        vendor_penalty = (
            df[df["status"] == "Delayed"]
            .groupby("vendor_code")["penalty"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        if not vendor_penalty.empty:
            fig = px.bar(
                vendor_penalty,
                x="vendor_code",
                y="penalty",
                title="Top 10 Vendors by Penalty",
                color="vendor_code",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_layout(xaxis_title="", yaxis_title="Penalty (₹)")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No delayed complaints with penalties.")

    # delay severity distribution
    st.subheader("Delay Severity")
    delayed_df = df[df["status"] == "Delayed"].copy()
    if not delayed_df.empty:
        delayed_df["severity"] = pd.cut(
            delayed_df["delay_hours"],
            bins=[0, 6, 24, 48, 72, 168, float("inf")],
            labels=["< 6 hrs", "6-24 hrs", "1-2 days", "2-3 days", "3-7 days", "7+ days"],
            right=False,
        )
        severity_counts = delayed_df["severity"].value_counts().reindex(
            ["< 6 hrs", "6-24 hrs", "1-2 days", "2-3 days", "3-7 days", "7+ days"]
        ).reset_index()
        severity_counts.columns = ["severity", "count"]
        fig = px.bar(
            severity_counts,
            x="severity",
            y="count",
            title="Delay Severity Distribution",
            color="severity",
            color_discrete_sequence=px.colors.sequential.Reds_r,
        )
        fig.update_layout(xaxis_title="", yaxis_title="Count")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No delayed complaints.")

    # ----------------------------------------------------------
    # data table with filters
    # ----------------------------------------------------------
    st.subheader("Complaint Records")

    status_filter = st.multiselect(
        "Filter by status",
        options=["Early", "Delayed", "On Time", "Pending"],
        default=[],
    )

    display_df = df.copy()
    if status_filter:
        display_df = display_df[display_df["status"].isin(status_filter)]

    # pick columns for display
    show_cols = [
        "complaint_id", "ro_name", "vendor_code", "assignment_time",
        "due_time", "close_time", "duration_text", "delay_hours",
        "status", "penalty", "is_auto_closed",
    ]

    table_df = display_df[show_cols].copy()

    # format datetimes for display
    for col in ["assignment_time", "due_time", "close_time"]:
        table_df[col] = table_df[col].apply(
            lambda x: x.strftime("%d-%b-%y %I:%M %p") if pd.notna(x) and x else "Pending"
        )

    table_df["penalty"] = table_df["penalty"].apply(lambda x: f"₹{x:,.0f}" if x else "₹0")
    table_df["is_auto_closed"] = table_df["is_auto_closed"].apply(
        lambda x: "⚠ Auto" if x else ""
    )

    # rename for readability
    table_df.columns = [
        "ID", "RO Name", "Vendor", "Assigned", "Due", "Closed",
        "Duration", "Hours", "Status", "Penalty", "Flag",
    ]

    styled = table_df.style.applymap(style_status, subset=["Status"])
    st.dataframe(styled, use_container_width=True, height=400)

    # ----------------------------------------------------------
    # export
    # ----------------------------------------------------------
    st.subheader("Export")

    export_col1, export_col2 = st.columns(2)

    with export_col1:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download CSV",
            data=csv,
            file_name="complaint_analysis.csv",
            mime="text/csv",
        )

    with export_col2:
        # create an excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Delay Analysis", index=False)
        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="complaint_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # auto-close warning
    auto_count = df["is_auto_closed"].sum()
    if auto_count > 0:
        st.warning(
            f"⚠ {auto_count} complaint(s) were auto-closed by the system "
            "(vendor didn't actually resolve). Check the 'Flag' column."
        )


if __name__ == "__main__":
    run()
