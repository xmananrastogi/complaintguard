"""
streamlit dashboard for the complaint sla analyser.
upload an excel, see results, drill down by nature / engineer / du.

run with:
    streamlit run app/main.py
"""

import os
import sys
import tempfile
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app.processor import (
    process_excel,
    analyse_natures,
    analyse_engineers,
    analyse_du_revisits,
    analyse_modes,
    generate_report_html,
)
from app.database import init_db, save_upload, get_upload_history

st.set_page_config(page_title="ComplaintGuard", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

COLORS = {"Early": "#c6efce", "Delayed": "#ffc7ce", "On Time": "#c6e0ff", "Pending": "#ffeb9c"}


def style_status(val: str) -> str:
    bg = COLORS.get(val, "white")
    return f"background-color: {bg}; font-weight: bold"


def show_summary_cards(summary):
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total", summary["total"])
    col2.metric("Early", summary["early"])
    col3.metric("Delayed", summary["delayed"])
    col4.metric("Pending", summary["pending"])
    col5.metric("Total Penalty", f"₹{summary['total_penalty']:,.0f}")


def run():
    init_db()
    st.title("🛡️ ComplaintGuard")
    st.markdown("SLA compliance analyser for IOCL vendor complaints.")

    with st.sidebar:
        st.header("Upload")
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "xlsm"])

        st.divider()
        st.header("History")
        history = get_upload_history()
        if history:
            for h in history[:5]:
                st.caption(f"{h['filename']} — {h['uploaded_at'][:10]}")
                st.caption(f"  {h['total_rows']} rows | ₹{h['total_penalty']:,.0f}")
        else:
            st.caption("No previous uploads yet.")

    if uploaded_file is None:
        st.info("Upload an Excel file to get started.")
        st.markdown("""
        **Expected columns:**
        - `Complaint ID`, `Complaint Resolution Time`, `Complaint DateTime`
        - `Vendor Close DateTime`, `Vendor Code`, `Vendor Remarks`
        - `RO Code`, `RO Name`, `Engineer Name`, `Nature of Complaint`
        - `DU serial No`, `Comp Mode`
        """)
        return

    # ----------------------------------------------------------
    # process
    # ----------------------------------------------------------
    with st.spinner("Processing complaints..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        try:
            records, summary = process_excel(tmp_path)
        except Exception as e:
            st.error(f"Failed to process file: {e}")
            os.unlink(tmp_path)
            return
        os.unlink(tmp_path)

    if not records:
        st.warning("No valid complaint records found.")
        return

    save_upload(uploaded_file.name, records, summary)
    df = pd.DataFrame(records)

    # compute aggregations
    natures = analyse_natures(records)
    engineers = analyse_engineers(records)
    revisits = analyse_du_revisits(records)
    modes = analyse_modes(records)

    # ----------------------------------------------------------
    # tabs
    # ----------------------------------------------------------
    tab_names = ["Overview", "Root Cause", "Engineers", "DU Revisits", "Comparison", "Report"]
    tabs = st.tabs(tab_names)

    # ===================== OVERVIEW =====================
    with tabs[0]:
        show_summary_cards(summary)

        st.subheader("Status Distribution")
        c1, c2 = st.columns(2)
        with c1:
            sc = df["status"].value_counts().reset_index()
            sc.columns = ["status", "count"]
            fig = px.pie(sc, values="count", names="status", color="status",
                         color_discrete_map=COLORS, title="Complaint Status")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            vp = df[df["status"] == "Delayed"].groupby("vendor_code")["penalty"].sum().sort_values(ascending=False).head(10).reset_index()
            if not vp.empty:
                fig = px.bar(vp, x="vendor_code", y="penalty", title="Top Vendors by Penalty",
                             color="vendor_code", color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(xaxis_title="", yaxis_title="Penalty (₹)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No delayed complaints.")

        st.subheader("Delay Severity")
        delayed_df = df[df["status"] == "Delayed"].copy()
        if not delayed_df.empty:
            delayed_df["sev"] = pd.cut(delayed_df["delay_hours"],
                bins=[0, 6, 24, 48, 72, 168, float("inf")],
                labels=["<6h", "6-24h", "1-2d", "2-3d", "3-7d", "7d+"], right=False)
            sv = delayed_df["sev"].value_counts().reindex(["<6h", "6-24h", "1-2d", "2-3d", "3-7d", "7d+"]).reset_index()
            sv.columns = ["sev", "count"]
            fig = px.bar(sv, x="sev", y="count", title="Delay Severity",
                         color="sev", color_discrete_sequence=px.colors.sequential.Reds_r)
            fig.update_layout(xaxis_title="", yaxis_title="Count")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No delayed complaints.")

        st.subheader("All Complaints")
        sf = st.multiselect("Filter by status", options=["Early", "Delayed", "On Time", "Pending"], default=[])
        td = df.copy()
        if sf:
            td = td[td["status"].isin(sf)]

        show_df = td[["complaint_id", "ro_name", "vendor_code", "assignment_time", "due_time",
                       "close_time", "duration_text", "delay_hours", "status", "penalty", "is_auto_closed"]].copy()
        for c in ["assignment_time", "due_time", "close_time"]:
            show_df[c] = show_df[c].apply(lambda x: x.strftime("%d-%b-%y %I:%M %p") if pd.notna(x) and x else "Pending")
        show_df["penalty"] = show_df["penalty"].apply(lambda x: f"₹{x:,.0f}")
        show_df["is_auto_closed"] = show_df["is_auto_closed"].apply(lambda x: "⚠ Auto" if x else "")
        show_df.columns = ["ID", "RO Name", "Vendor", "Assigned", "Due", "Closed", "Duration", "Hours", "Status", "Penalty", "Flag"]
        styled = show_df.style.map(style_status, subset=["Status"])
        st.dataframe(styled, use_container_width=True, height=400)

        ac = df["is_auto_closed"].sum()
        if ac:
            st.warning(f"⚠ {ac} complaints were auto-closed (vendor didn't actually resolve).")

    # ===================== ROOT CAUSE =====================
    with tabs[1]:
        st.subheader("Complaint Type Analysis")
        st.caption("Which nature of complaint causes the most delays.")

        ndf = pd.DataFrame(natures)
        c1, c2 = st.columns(2)
        with c1:
            topn = ndf.head(15)
            fig = px.bar(topn, x="nature", y="delay_rate", title="Delay Rate by Complaint Type",
                         color="delay_rate", color_continuous_scale="Reds", text_auto=".1f")
            fig.update_layout(xaxis_title="", yaxis_title="Delay %")
            fig.update_xaxes(tickangle=45)
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig = px.bar(topn, x="nature", y="avg_delay_hours", title="Avg Delay Hours by Type",
                         color="avg_delay_hours", color_continuous_scale="Oranges", text_auto=".1f")
            fig.update_layout(xaxis_title="", yaxis_title="Avg Hours")
            fig.update_xaxes(tickangle=45)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Detailed Table")
        ndf["penalty"] = ndf["penalty"].apply(lambda x: f"₹{x:,.0f}")
        ndf.columns = ["Nature", "Total", "Delayed", "Early", "Delay%", "Avg Delay(h)", "Penalty"]
        st.dataframe(ndf, use_container_width=True)

        st.divider()
        st.subheader("Complaint Mode (WEB vs SYSTEM)")
        mdf = pd.DataFrame(modes)
        if not mdf.empty:
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(mdf, values="total", names="mode", title="Complaints by Mode",
                             color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(mdf, x="mode", y="delay_rate", title="Delay Rate by Mode",
                             color="mode", text_auto=".1f", color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(xaxis_title="", yaxis_title="Delay %")
                st.plotly_chart(fig, use_container_width=True)
            mdf["penalty"] = mdf["penalty"].apply(lambda x: f"₹{x:,.0f}")
            mdf.columns = ["Mode", "Total", "Delayed", "Delay%", "Penalty"]
            st.dataframe(mdf, use_container_width=True)

    # ===================== ENGINEERS =====================
    with tabs[2]:
        st.subheader("Engineer Performance")
        st.caption("Ranked by compliance rate (worst first).")

        edf = pd.DataFrame(engineers)
        if not edf.empty:
            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(edf.head(20), x="engineer", y="compliance_rate",
                             title="Compliance Rate (worst → best)",
                             color="compliance_rate", color_continuous_scale="RdYlGn", text_auto=".1f")
                fig.update_layout(xaxis_title="", yaxis_title="Compliance %")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                fig = px.bar(edf.head(20), x="engineer", y="total",
                             title="Complaints Handled per Engineer",
                             color="total", color_continuous_scale="Blues", text_auto=".0f")
                fig.update_layout(xaxis_title="", yaxis_title="Count")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Engineer Scorecard")
            edf_display = edf.copy()
            edf_display["penalty"] = edf_display["penalty"].apply(lambda x: f"₹{x:,.0f}")
            edf_display.columns = ["Engineer", "Total", "Delayed", "Early", "On Time", "Compliance%", "Avg Delay(h)", "Penalty"]
            st.dataframe(edf_display, use_container_width=True)

            # highlight bottom 3
            worst = edf.head(3)
            st.warning("🔴 Bottom 3 performers:")
            for _, r in worst.iterrows():
                st.write(f"  **{r['engineer']}** — {r['delayed']} delays out of {r['total']} ({100 - r['compliance_rate']:.1f}% delay rate)")
        else:
            st.info("No engineer data found.")

    # ===================== DU REVISITS =====================
    with tabs[3]:
        st.subheader("DU Revisit Tracker")
        st.caption("Same equipment flagged multiple times within 30 days = incomplete fix.")

        if revisits:
            rv_summary = []
            for rv in revisits:
                rv_summary.append({
                    "DU Serial": rv["du_serial"],
                    "Complaints": rv["total_complaints"],
                    "Revisits": rv["revisit_count"],
                    "Vendor": rv.get("vendor", ""),
                    "RO Name": rv.get("ro_name", ""),
                })
            rvdf = pd.DataFrame(rv_summary)
            st.dataframe(rvdf, use_container_width=True)

            st.subheader("Revisit Details")
            selected_du = st.selectbox("Select DU to inspect", [rv["du_serial"] for rv in revisits])
            for rv in revisits:
                if rv["du_serial"] == selected_du:
                    details = []
                    for rev in rv["revisits"]:
                        details.append({
                            "Previous ID": rev["prev_id"],
                            "Next ID": rev["next_id"],
                            "Gap (days)": rev["gap_days"],
                            "Prev Status": rev["prev_status"],
                            "Next Status": rev["next_status"],
                        })
                    st.dataframe(pd.DataFrame(details), use_container_width=True)
                    break
        else:
            st.success("No DU revisits found within 30-day window.")

    # ===================== COMPARISON =====================
    with tabs[4]:
        st.subheader("Compare Two Files")
        st.caption("Upload a second file to see how stats changed.")

        file2 = st.file_uploader("Upload second file for comparison", type=["xlsx", "xls", "xlsm"], key="comp")
        if file2:
            with st.spinner("Processing second file..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
                    tmp.write(file2.getvalue())
                    tmp_path = tmp.name
                try:
                    recs2, sum2 = process_excel(tmp_path)
                except:
                    st.error("Failed to process second file.")
                    os.unlink(tmp_path)
                    st.stop()
                os.unlink(tmp_path)

            if recs2:
                st.subheader("Side-by-Side Comparison")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**File 1 (current)**")
                    st.json(summary)
                with c2:
                    st.markdown("**File 2 (comparison)**")
                    st.json(sum2)

                # delta table
                delta = {
                    "Metric": ["Total", "Early", "Delayed", "On Time", "Pending", "Penalty"],
                    "File 1": [summary["total"], summary["early"], summary["delayed"],
                               summary["on_time"], summary["pending"], f"₹{summary['total_penalty']:,.0f}"],
                    "File 2": [sum2["total"], sum2["early"], sum2["delayed"],
                               sum2["on_time"], sum2["pending"], f"₹{sum2['total_penalty']:,.0f}"],
                    "Change": [
                        sum2["total"] - summary["total"],
                        sum2["early"] - summary["early"],
                        sum2["delayed"] - summary["delayed"],
                        sum2["on_time"] - summary["on_time"],
                        sum2["pending"] - summary["pending"],
                        f"₹{sum2['total_penalty'] - summary['total_penalty']:,.0f}",
                    ],
                }
                st.dataframe(pd.DataFrame(delta), use_container_width=True)

    # ===================== REPORT =====================
    with tabs[5]:
        st.subheader("Generate Report")
        st.caption("Preview and download a printable HTML report.")

        html = generate_report_html(records, summary)
        st.components.v1.html(html, height=500, scrolling=True)

        st.download_button("Download Report (HTML)", data=html, file_name="complaint_report.html", mime="text/html")

        # also csv / excel export
        c1, c2 = st.columns(2)
        with c1:
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button("Download CSV", data=csv, file_name="complaint_analysis.csv", mime="text/csv")
        with c2:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Delay Analysis", index=False)
            st.download_button("Download Excel", data=output.getvalue(),
                               file_name="complaint_analysis.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    run()
