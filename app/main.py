"""
streamlit dashboard for the complaint sla analyser.
upload an excel, see results, drill down by nature / engineer / vendor / ro / du.

run with:
    streamlit run app/main.py
"""

import os
import sys
import tempfile
from io import BytesIO
from datetime import datetime, timedelta
from copy import deepcopy

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
    analyse_vendors,
    analyse_ros,
    generate_report_html,
)
from app.database import init_db, save_upload, get_upload_history
from app.config_loader import penalty_rate as cfg_penalty_rate, revisit_window as cfg_revisit_window, load as cfg_load

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


def format_time(dt):
    if pd.isna(dt) or not dt:
        return "Pending"
    return dt.strftime("%d-%b-%y %I:%M %p")


def show_complaint_table(df_subset, key_suffix=""):
    """render a styled complaint table with drill-down rows."""
    if df_subset.empty:
        st.info("No matching complaints.")
        return

    show = df_subset[["complaint_id", "ro_name", "vendor_code", "assignment_time", "due_time",
                       "close_time", "duration_text", "delay_hours", "status", "penalty", "is_auto_closed"]].copy()
    for c in ["assignment_time", "due_time", "close_time"]:
        show[c] = show[c].apply(format_time)
    show["penalty"] = show["penalty"].apply(lambda x: f"₹{x:,.0f}")
    show["is_auto_closed"] = show["is_auto_closed"].apply(lambda x: "⚠ Auto" if x else "")
    show.columns = ["ID", "RO Name", "Vendor", "Assigned", "Due", "Closed", "Duration", "Hours", "Status", "Penalty", "Flag"]
    styled = show.style.map(style_status, subset=["Status"])
    st.dataframe(styled, use_container_width=True, height=400)


def export_data(df, filename_base):
    """csv + excel download buttons in a 2-column layout."""
    c1, c2 = st.columns(2)
    with c1:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            f"📥 Download CSV ({len(df)} rows)",
            data=csv,
            file_name=f"{filename_base}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c2:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Delay Analysis", index=False)
        st.download_button(
            f"📥 Download Excel ({len(df)} rows)",
            data=output.getvalue(),
            file_name=f"{filename_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def build_table(records):
    """convert a list of record dicts into a display dataframe."""
    return pd.DataFrame(records)


def apply_penalty_rate(records, rate):
    """recompute penalties and summary for a given per-block rate."""
    out = []
    s = {"total": 0, "early": 0, "delayed": 0, "on_time": 0, "pending": 0, "total_penalty": 0.0}
    for r in records:
        row = dict(r)
        status = r["status"]
        s["total"] += 1
        if status == "Early":
            s["early"] += 1
            row["penalty"] = 0
        elif status == "Delayed":
            s["delayed"] += 1
            row["penalty"] = int(r["delay_hours"] / 24) * rate
            s["total_penalty"] += row["penalty"]
        elif status == "On Time":
            s["on_time"] += 1
            row["penalty"] = 0
        else:
            s["pending"] += 1
            row["penalty"] = 0
        out.append(row)
    return out, s


def run():
    init_db()
    st.title("🛡️ ComplaintGuard")
    st.markdown("SLA compliance analyser for IOCL vendor complaints.")

    # ----------- session state -----------
    if "records" not in st.session_state:
        st.session_state.records = None
    if "summary" not in st.session_state:
        st.session_state.summary = None
    if "col_map" not in st.session_state:
        st.session_state.col_map = None
    if "penalty_rate" not in st.session_state:
        st.session_state.penalty_rate = cfg_penalty_rate()
    if "revisit_window" not in st.session_state:
        st.session_state.revisit_window = cfg_revisit_window()

    # ----------- sidebar -----------
    with st.sidebar:
        st.header("Upload")
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "xlsm"])

        st.divider()

        # settings
        with st.expander("⚙️ Settings", expanded=False):
            st.caption("Values from `config/settings.yaml`. Override for this session below.")
            rate = st.number_input(
                "Penalty per 24h block (₹)",
                min_value=0, max_value=10000, value=int(st.session_state.penalty_rate), step=100,
                help="Amount charged for each full 24-hour delay block.",
            )
            if rate != st.session_state.penalty_rate and st.session_state.records is not None:
                st.session_state.penalty_rate = rate
                st.session_state.records, st.session_state.summary = apply_penalty_rate(
                    st.session_state._orig_records, rate
                )
                st.rerun()

            revisit = st.number_input(
                "DU revisit window (days)",
                min_value=1, max_value=365, value=int(st.session_state.revisit_window), step=1,
                help="Days within which same-DU repeat counts as a revisit.",
            )
            if revisit != st.session_state.revisit_window:
                st.session_state.revisit_window = revisit
                st.rerun()

        st.divider()

        # export — only if data loaded
        if st.session_state.records is not None:
            st.header("Export")
            export_data(pd.DataFrame(st.session_state.records), "complaint_analysis")
            st.divider()

        st.header("History")
        history = get_upload_history()
        if history:
            for h in history[:5]:
                st.caption(f"{h['filename']} — {h['uploaded_at'][:10]}")
                st.caption(f"  {h['total_rows']} rows | ₹{h['total_penalty']:,.0f}")
        else:
            st.caption("No previous uploads yet.")

    # ----------- no upload state -----------
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

    # ----------- process on new upload -----------
    if st.session_state.records is None or st.session_state.get("_uploaded_name") != uploaded_file.name:
        with st.spinner("Processing complaints..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            try:
                records, summary, col_map = process_excel(tmp_path, penalty_per_block=rate)
            except Exception as e:
                st.error(f"Failed to process file: {e}")
                os.unlink(tmp_path)
                return
            os.unlink(tmp_path)

        if not records:
            st.warning("No valid complaint records found.")
            if "error" in summary:
                st.error(summary["error"])
            return

        st.session_state._orig_records = records
        st.session_state.penalty_rate = rate
        st.session_state.records, st.session_state.summary = apply_penalty_rate(records, rate)
        st.session_state.col_map = col_map
        st.session_state._uploaded_name = uploaded_file.name

        save_upload(uploaded_file.name, st.session_state.records, st.session_state.summary)

    # ----------- data ready -----------
    records = st.session_state.records
    summary = st.session_state.summary
    col_map = st.session_state.col_map
    df = pd.DataFrame(records)

    # show column mapping
    missing_cols = {k: v for k, v in col_map.items() if v is None}
    if missing_cols:
        st.warning(f"Could not detect: {', '.join(missing_cols.keys())}. Check column names.")
    with st.expander("Detected Column Mapping", expanded=False):
        cols = pd.DataFrame([
            {"Field": k, "File Column": v or "❌ Not found"}
            for k, v in col_map.items()
        ])
        st.dataframe(cols, use_container_width=True, hide_index=True)

    # ----------- sidebar filters -----------
    with st.sidebar:
        st.header("Filters")

        date_col = "assignment_time"
        if date_col in df.columns and df[date_col].notna().any():
            valid_dates = df[df[date_col].notna()][date_col]
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            date_range = st.date_input("Date range", [min_date, max_date], min_value=min_date, max_value=max_date)
        else:
            date_range = [None, None]

        status_opts = ["Early", "Delayed", "On Time", "Pending"]
        sel_status = st.multiselect("Status", status_opts, default=[])

        ro_opts = sorted(df["ro_name"].dropna().unique()) if "ro_name" in df.columns else []
        sel_ro = st.multiselect("RO Name", ro_opts, default=[])

        vendor_opts = sorted(df["vendor_code"].dropna().unique()) if "vendor_code" in df.columns else []
        sel_vendor = st.multiselect("Vendor", vendor_opts, default=[])

        engineer_opts = sorted(df["engineer_name"].dropna().unique()) if "engineer_name" in df.columns else []
        sel_engineer = st.multiselect("Engineer", engineer_opts, default=[])

    # ----------- apply filters -----------
    filtered = df.copy()
    if date_range[0] and date_col in filtered.columns:
        lo, hi = date_range[0], date_range[1] if len(date_range) > 1 else date_range[0]
        filtered = filtered[
            filtered[date_col].notna()
            & (filtered[date_col].dt.date >= lo)
            & (filtered[date_col].dt.date <= hi + timedelta(days=1))
        ]
    if sel_status:
        filtered = filtered[filtered["status"].isin(sel_status)]
    if sel_ro:
        filtered = filtered[filtered["ro_name"].isin(sel_ro)]
    if sel_vendor:
        filtered = filtered[filtered["vendor_code"].isin(sel_vendor)]
    if sel_engineer:
        filtered = filtered[filtered["engineer_name"].isin(sel_engineer)]

    filtered_records = filtered.to_dict("records")

    # ----------- compute aggregations on filtered data -----------
    natures = analyse_natures(filtered_records)
    engineers = analyse_engineers(filtered_records)
    revisits = analyse_du_revisits(filtered_records, window_days=st.session_state.revisit_window)
    modes = analyse_modes(filtered_records)
    vendors = analyse_vendors(filtered_records)
    ros = analyse_ros(filtered_records)

    filter_active = len(filtered) != len(df)
    if filter_active:
        st.info(f"Showing {len(filtered)} of {len(df)} complaints (filters active).")

    # ----------- tabs -----------
    tab_names = ["Overview", "Root Cause", "Engineers", "Vendors", "ROs", "DU Revisits", "Comparison", "Report"]
    tabs = st.tabs(tab_names)

    # ===================== OVERVIEW =====================
    with tabs[0]:
        show_summary_cards(summary)

        st.subheader("Status Distribution")
        c1, c2 = st.columns(2)
        with c1:
            sc = filtered["status"].value_counts().reset_index()
            sc.columns = ["status", "count"]
            fig = px.pie(sc, values="count", names="status", color="status",
                         color_discrete_map=COLORS, title="Complaint Status")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            vp = filtered[filtered["status"] == "Delayed"].groupby("vendor_code")["penalty"].sum().sort_values(ascending=False).head(10).reset_index()
            if not vp.empty:
                fig = px.bar(vp, x="vendor_code", y="penalty", title="Top Vendors by Penalty",
                             color="vendor_code", color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(xaxis_title="", yaxis_title="Penalty (₹)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No delayed complaints.")

        st.subheader("Delay Severity")
        delayed_df = filtered[filtered["status"] == "Delayed"].copy()
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
        show_complaint_table(filtered)

        ac = df["is_auto_closed"].sum()
        if ac:
            st.warning(f"⚠ {ac} complaints were auto-closed (vendor didn't actually resolve).")

    # ===================== ROOT CAUSE =====================
    with tabs[1]:
        st.subheader("Complaint Type Analysis")
        st.caption("Which nature of complaint causes the most delays. Select one below to drill into individual complaints.")

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
        ndf_display = ndf.copy()
        ndf_display["penalty"] = ndf_display["penalty"].apply(lambda x: f"₹{x:,.0f}")
        ndf_display.columns = ["Nature", "Total", "Delayed", "Early", "Delay%", "Avg Delay(h)", "Penalty"]
        st.dataframe(ndf_display, use_container_width=True)

        # drill-down
        if not ndf.empty:
            sel_nature = st.selectbox("🔍 View complaints by nature", [""] + list(ndf["nature"]), key="drill_nature")
            if sel_nature:
                sub = filtered[filtered["nature"] == sel_nature]
                show_complaint_table(sub)

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
        st.caption("Ranked by compliance rate (worst first). Select one below to drill into individual complaints.")

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
                             title="Complaints Handled",
                             color="total", color_continuous_scale="Blues", text_auto=".0f")
                fig.update_layout(xaxis_title="", yaxis_title="Count")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Engineer Scorecard")
            edf_display = edf.copy()
            edf_display["penalty"] = edf_display["penalty"].apply(lambda x: f"₹{x:,.0f}")
            edf_display.columns = ["Engineer", "Total", "Delayed", "Early", "On Time", "Compliance%", "Avg Delay(h)", "Penalty"]
            st.dataframe(edf_display, use_container_width=True)

            # drill-down
            sel_eng = st.selectbox("🔍 View complaints by engineer", [""] + list(edf["engineer"]), key="drill_eng")
            if sel_eng:
                sub = filtered[filtered["engineer_name"] == sel_eng]
                show_complaint_table(sub)

            # bottom 3
            worst = edf.head(3)
            st.warning("🔴 Bottom 3 performers:")
            for _, r in worst.iterrows():
                st.write(f"  **{r['engineer']}** — {r['delayed']} delays out of {r['total']} ({100 - r['compliance_rate']:.1f}% delay rate)")
        else:
            st.info("No engineer data found.")

    # ===================== VENDORS =====================
    with tabs[3]:
        st.subheader("Vendor Performance")
        st.caption("Who's racking up the most penalties. Select a vendor to drill into individual complaints.")

        vdf = pd.DataFrame(vendors)
        if not vdf.empty:
            c1, c2 = st.columns(2)
            with c1:
                topv = vdf.head(15)
                fig = px.bar(topv, x="vendor", y="penalty", title="Penalty by Vendor (top 15)",
                             color="penalty", color_continuous_scale="Reds", text_auto=".0s")
                fig.update_layout(xaxis_title="", yaxis_title="Penalty (₹)")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                fig = px.bar(topv, x="vendor", y="compliance_rate", title="Compliance Rate by Vendor",
                             color="compliance_rate", color_continuous_scale="RdYlGn", text_auto=".1f")
                fig.update_layout(xaxis_title="", yaxis_title="Compliance %")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Vendor Scorecard")
            vdf_display = vdf.copy()
            vdf_display["penalty"] = vdf_display["penalty"].apply(lambda x: f"₹{x:,.0f}")
            vdf_display.columns = ["Vendor", "Total", "Delayed", "Early", "On Time", "Pending", "Compliance%", "Penalty"]
            st.dataframe(vdf_display, use_container_width=True)

            # drill-down
            sel_v = st.selectbox("🔍 View complaints by vendor", [""] + list(vdf["vendor"]), key="drill_vendor")
            if sel_v:
                sub = filtered[filtered["vendor_code"] == sel_v]
                show_complaint_table(sub)
        else:
            st.info("No vendor data found.")

    # ===================== ROs =====================
    with tabs[4]:
        st.subheader("Retail Outlet Issues")
        st.caption("Which ROs have the most complaints and delays. Select one to drill in.")

        rdf = pd.DataFrame(ros)
        if not rdf.empty:
            c1, c2 = st.columns(2)
            with c1:
                top_ro = rdf.head(15)
                fig = px.bar(top_ro, x="ro_code", y="total", title="Complaints by RO (top 15)",
                             color="total", color_continuous_scale="Blues", text_auto=".0f",
                             hover_data=["ro_name"])
                fig.update_layout(xaxis_title="", yaxis_title="Complaints")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                fig = px.bar(top_ro, x="ro_code", y="penalty", title="Penalty by RO",
                             color="penalty", color_continuous_scale="Reds", text_auto=".0s",
                             hover_data=["ro_name"])
                fig.update_layout(xaxis_title="", yaxis_title="Penalty (₹)")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("RO Scorecard")
            rdf_display = rdf.copy()
            rdf_display["penalty"] = rdf_display["penalty"].apply(lambda x: f"₹{x:,.0f}")
            rdf_display.columns = ["RO Code", "RO Name", "Total", "Delayed", "Delay%", "Penalty"]
            st.dataframe(rdf_display, use_container_width=True)

            # drill-down
            sel_ro_dd = st.selectbox("🔍 View complaints by RO", [""] + list(rdf["ro_code"]), key="drill_ro")
            if sel_ro_dd:
                sub = filtered[filtered["ro_code"] == sel_ro_dd]
                show_complaint_table(sub)
        else:
            st.info("No RO data found.")

    # ===================== DU REVISITS =====================
    with tabs[5]:
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
    with tabs[6]:
        st.subheader("Compare Two Files")
        st.caption("Upload a second file to see how stats changed side-by-side.")

        file2 = st.file_uploader("Upload second file for comparison", type=["xlsx", "xls", "xlsm"], key="comp")
        if file2:
            with st.spinner("Processing second file..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
                    tmp.write(file2.getvalue())
                    tmp_path = tmp.name
                try:
                    recs2, sum2, _ = process_excel(tmp_path, penalty_per_block=rate)
                except Exception as e:
                    st.error(f"Failed to process second file: {e}")
                    os.unlink(tmp_path)
                    st.stop()
                os.unlink(tmp_path)

            if recs2:
                # counts
                df2 = pd.DataFrame(recs2)

                st.subheader("Summary Comparison")
                c1, c2, c3 = st.columns(3)
                c1.metric("Total", sum2["total"], delta=sum2["total"] - summary["total"])
                c2.metric("Delayed", sum2["delayed"], delta=sum2["delayed"] - summary["delayed"])
                c3.metric("Penalty", f"₹{sum2['total_penalty']:,.0f}",
                          delta=f"₹{sum2['total_penalty'] - summary['total_penalty']:,.0f}")

                # side-by-side status chart
                comp_df = pd.DataFrame({
                    "Status": ["Early", "Delayed", "On Time", "Pending"],
                    "File 1": [summary["early"], summary["delayed"], summary["on_time"], summary["pending"]],
                    "File 2": [sum2["early"], sum2["delayed"], sum2["on_time"], sum2["pending"]],
                }).melt(id_vars="Status", var_name="File", value_name="Count")

                fig = px.bar(comp_df, x="Status", y="Count", color="File", barmode="group",
                             title="Status Distribution Comparison",
                             color_discrete_sequence=["#3498db", "#e74c3c"])
                st.plotly_chart(fig, use_container_width=True)

                # top vendors comparison
                v1 = pd.DataFrame(analyse_vendors(recs2))
                v2 = pd.DataFrame(vendors)
                if not v1.empty and not v2.empty:
                    merged = v2.merge(v1, on="vendor", how="outer", suffixes=("_1", "_2")).fillna(0)
                    merged = merged.head(10)
                    fig = go.Figure()
                    fig.add_trace(go.Bar(name="File 1", x=merged["vendor"], y=merged["penalty_1"], marker_color="#3498db"))
                    fig.add_trace(go.Bar(name="File 2", x=merged["vendor"], y=merged["penalty_2"], marker_color="#e74c3c"))
                    fig.update_layout(title="Penalty by Vendor — File 1 vs File 2", barmode="group",
                                      xaxis_title="", yaxis_title="Penalty (₹)")
                    st.plotly_chart(fig, use_container_width=True)

                # full delta table
                st.subheader("Metric Comparison")
                delta = {
                    "Metric": ["Total", "Early", "Delayed", "On Time", "Pending", "Penalty"],
                    "File 1": [summary["total"], summary["early"], summary["delayed"],
                               summary["on_time"], summary["pending"], f"₹{summary['total_penalty']:,.0f}"],
                    "File 2": [sum2["total"], sum2["early"], sum2["delayed"],
                               sum2["on_time"], sum2["pending"], f"₹{sum2['total_penalty']:,.0f}"],
                    "Δ": [sum2["total"] - summary["total"],
                          sum2["early"] - summary["early"],
                          sum2["delayed"] - summary["delayed"],
                          sum2["on_time"] - summary["on_time"],
                          sum2["pending"] - summary["pending"],
                          f"₹{sum2['total_penalty'] - summary['total_penalty']:,.0f}"],
                }
                st.dataframe(pd.DataFrame(delta), use_container_width=True)

    # ===================== REPORT =====================
    with tabs[7]:
        st.subheader("Generate Report")
        st.caption("Preview and download a printable HTML report.")

        html = generate_report_html(records, summary)
        st.components.v1.html(html, height=500, scrolling=True)

        st.download_button("📄 Download Report (HTML)", data=html, file_name="complaint_report.html", mime="text/html", use_container_width=True)


if __name__ == "__main__":
    run()
