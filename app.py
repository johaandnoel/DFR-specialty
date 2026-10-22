from datetime import datetime
from pathlib import Path
from time import perf_counter
from zoneinfo import ZoneInfo

import streamlit as st

from jobinfo import JOBINFO_URL, candidate_job_ids, fetch_live_jobs, verified_excel_rows
from job_reports import canonical_job_id, rank_references
from reference_search import search_workbook


st.set_page_config(page_title="Closest reference search", page_icon=":material/compare:", layout="wide")
sample = Path(__file__).with_name("sample daily tracker.xlsx")

with st.sidebar:
    st.subheader("Reference source", anchor=False)
    st.caption("Daily tracker")
    uploaded = st.file_uploader("Replace tracker Excel", type=["xlsx"])
    st.caption(f"Using: {uploaded.name if uploaded else sample.name}")
    st.divider()
    st.markdown("**Live reports only**")
    st.caption("Job Text and BOM By Operation are retrieved from JobInfo using your Windows login.")
    st.link_button("Open JobInfo", JOBINFO_URL, icon=":material/open_in_new:", width="stretch")
    with st.expander("How matching works"):
        st.caption(
            "Item description uses a contains match; drawing number uses an exact match. "
            "Both must match when provided. PART rows are excluded."
        )
        st.caption(
            "Master Job values are lookup leads. The returned job ID, item, and drawing "
            "must agree with the Excel row before a reference is ranked."
        )
        st.caption(
            "Ranking: exact matches, complete coil/fan evidence, matching coil/fan components, "
            "job-text resemblance, BOM overlap, then valid completion date. "
            "Percentages are comparison metrics, not release confidence."
        )

st.title("Closest reference search", anchor=False)

with st.form("reference_search", border=True):
    st.subheader("Search a job", anchor=False)
    drawing_col, description_col, job_col = st.columns([1, 1.5, 1], vertical_alignment="bottom")
    drawing = drawing_col.text_input("Drawing number", placeholder="e.g. 160754", key="drawing")
    description = description_col.text_input(
        "Item description", placeholder="e.g. PF-4R-48-CBOM", key="description"
    )
    job_id = job_col.text_input("Job ID", placeholder="e.g. 322784", key="job_id")
    st.caption("Filter Excel by item and/or drawing. Enter a Job ID to compare live reports.")
    submitted = st.form_submit_button(
        "Find closest reference", type="primary", icon=":material/search:", width="stretch"
    )

if submitted:
    st.session_state.pop("search_result", None)
    started = perf_counter()
    status = st.status("Finding references...", expanded=True)
    progress_bar = st.progress(0, text="Filtering the daily tracker")
    activity = st.empty()
    try:
        matches = search_workbook(
            uploaded.getvalue() if uploaded else sample.read_bytes(),
            description, drawing, today=datetime.now(ZoneInfo("Asia/Kolkata")).date(),
        )
        result = {
            "matches": matches, "drawing": drawing.strip(), "description": description.strip(),
            "job_id": job_id.strip(), "source": uploaded.name if uploaded else sample.name,
        }
        st.session_state.search_result = result
        if job_id.strip():
            result.update(
                comparison_source="JobInfo", ranked=[], reports={}, retrieval_issues=[],
                excluded_rows=[], unresolved_rows=[],
            )
            target_id = canonical_job_id(job_id)
            reference_ids, result["unresolved_rows"] = candidate_job_ids(matches, target_id)
            result["reference_ids"] = reference_ids
            progress_bar.progress(0.1, text=f"{len(matches)} Excel rows · {len(reference_ids)} unique reference jobs")

            def report_progress(done, total, current_id, stage="Reading job reports"):
                progress_bar.progress(
                    0.1 + 0.8 * (done / total if total else 0),
                    text=f"{done}/{total} report steps processed",
                )
                role = "Requested job" if current_id == target_id else "Reference"
                activity.caption(f"{role} {current_id} — {stage}" if current_id else stage)
                status.update(
                    label=f"{stage} · {current_id}" if current_id else stage,
                    state="running",
                )

            jobs, result["retrieval_issues"] = fetch_live_jobs(
                [target_id, *reference_ids], progress=report_progress
            )
            result["reports"] = jobs
            progress_bar.progress(0.92, text="Verifying reference IDs, items, and drawings")
            verified, result["excluded_rows"] = verified_excel_rows(matches, jobs)
            progress_bar.progress(0.96, text="Comparing job text, coils, fans, and BOMs")
            status.update(label="Comparing specifications and ranking references...")
            result["ranked"], blocked = rank_references(target_id, verified, jobs)
            result["retrieval_issues"].extend(blocked)
            result["retrieved_at"] = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y, %H:%M IST")
            partial = bool(result["retrieval_issues"] or result["excluded_rows"] or result["unresolved_rows"])
            label = "Comparison complete — some references excluded" if partial else "Comparison complete"
        else:
            label = "Excel shortlist ready — enter a Job ID for live comparison"
        result["elapsed_seconds"] = round(perf_counter() - started, 1)
        progress_bar.progress(1.0, text=f"Finished in {result['elapsed_seconds']} seconds")
        activity.empty()
        status.update(label=label, state="complete", expanded=False)
    except Exception as exc:
        message = str(exc).splitlines()[0] or type(exc).__name__
        result = st.session_state.get("search_result")
        if result is not None:
            result["comparison_error"] = message
        activity.caption(message)
        status.update(label="Search could not be completed", state="error", expanded=True)
        if result is None:
            st.error(message)

result = st.session_state.get("search_result")
if result is None:
    with st.container(border=True):
        st.subheader("Start with the job you want to build", anchor=False)
        st.caption(
            "1. Enter the job ID and Excel filters.  "
            "2. We read live Job Text and BOM reports.  "
            "3. Review the closest reference and its differences."
        )
else:
    st.divider()
    st.caption(
        f"Requested job: {result['job_id'] or 'not entered'} · "
        f"Item: {result['description'] or 'any'} · Drawing: {result['drawing'] or 'any'}"
    )
    if result.get("comparison_error"):
        st.error(f"Live comparison unavailable: {result['comparison_error']}")
    ranked = result.get("ranked", []) if result.get("comparison_source") == "JobInfo" else []
    issues = result.get("retrieval_issues", []) + result.get("excluded_rows", [])
    invalid_dates = sum(row["Date status"] != "Valid" for row in result["matches"])
    if invalid_dates:
        issues.append(f"{invalid_dates} Excel rows have questionable completion dates and are listed last.")
    if result.get("unresolved_rows"):
        issues.append(
            f"{len(result['unresolved_rows'])} rows have no usable Master Job ID; "
            "they were not looked up. Drawing numbers are not job IDs."
        )
    if issues:
        st.warning("Some records could not be used. See search coverage below.", icon=":material/info:")

    if ranked:
        best = ranked[0]
        with st.container(border=True):
            st.caption("TOP VERIFIED MATCH")
            st.subheader(f"Reference {best['Job ID']}", anchor=False)
            if best["Recommendation"].startswith("Insufficient"):
                st.warning("Insufficient coil/fan evidence to confirm a closest reference.")
            elif best["Recommendation"].startswith("Potential copy"):
                st.success(f"Closest verified reference: {best['Job ID']} · Potential copy")
            else:
                st.success(f"Closest verified reference: {best['Job ID']}")
                st.markdown("**Use as a model skeleton — changes required.**")
            st.caption(
                f"{len(ranked)} verified references compared · "
                f"Retrieved {result.get('retrieved_at', 'from JobInfo')}"
            )
        with st.container(horizontal=True):
            st.metric("Job-text resemblance", f"{best['Job-text resemblance (%)']}%", border=True)
            st.metric("Exact BOM overlap", f"{best['Exact BOM-row overlap (%)']}%", border=True)
            st.metric("Coil components", "Match" if best["Coil component match"] else "Review", border=True)
            st.metric("Fan components", "Match" if best["Fan component match"] else "Review", border=True)

        st.subheader("Compare references", anchor=False)
        table = [{
            "Job ID": row["Job ID"], "Recommendation": row["Recommendation"],
            "Job text (%)": row["Job-text resemblance (%)"],
            "BOM overlap (%)": row["Exact BOM-row overlap (%)"],
            "Coil match": row["Coil component match"], "Fan match": row["Fan component match"],
            "Completed": row["Completed date"] or "Unknown",
        } for row in ranked]
        st.dataframe(table, hide_index=True, height=240, alt="Ranked live reference jobs")
        selected = st.selectbox("Inspect reference differences", [row["Job ID"] for row in ranked])
        comparison = next(row for row in ranked if row["Job ID"] == selected)
        target_key = next(
            key for key in result["reports"]
            if canonical_job_id(key) == canonical_job_id(result["job_id"])
        )
        target = result["reports"][target_key]["BOM"]
        reference = result["reports"][selected]["BOM"]
        with st.container(border=True):
            st.markdown(f"**{comparison['Recommendation']}**")
            st.subheader("Coils & fans", anchor=False)
            evidence = []
            for label, bom in (("Requested job", target), ("Reference job", reference)):
                for row in bom["coils"] + bom["fans"]:
                    evidence.append({
                        "Job": label, "Part": row["part"], "Size": row.get("size", "—"),
                        "Quantity": row["quantity"], "Operation": row["operation"],
                        "Description": row["description"],
                    })
            st.dataframe(evidence, hide_index=True, alt="Requested and reference coil and fan evidence")
            left, right = st.columns(2)
            left.metric("Requested job fan motor quantity", target["fan_count"] or "Unknown")
            right.metric("Reference fan motor quantity", reference["fan_count"] or "Unknown")
        with st.expander(f"Job-text differences ({len(comparison['section_changes'])} sections)"):
            if comparison["section_changes"]:
                st.dataframe(
                    comparison["section_changes"], hide_index=True, height=320,
                    alt="Changed job-text specification sections",
                )
            else:
                st.caption("No differences found in the compared job-text sections.")
        st.caption("Engineering review of drawings and model revisions is required before any copy release.")
    elif not result.get("comparison_error"):
        st.info(
            "No verified reference is available for this search. Adjust the filters or review coverage."
            if result.get("comparison_source") == "JobInfo"
            else "Excel shortlist ready. Enter a Job ID to compare live JobInfo reports."
        )

    with st.expander(f"Excel shortlist ({len(result['matches'])} rows)"):
        st.caption(f"Source: {result['source']} · Valid completion dates, newest first · Repeated tracker rows retained")
        if result["matches"]:
            st.dataframe(
                result["matches"], hide_index=True, height=260,
                alt="Filtered daily tracker candidate records",
            )
        else:
            st.caption("No rows match. Try a shorter description or omit the drawing number.")
    with st.expander("Search coverage & excluded records"):
        if issues:
            for issue in issues:
                st.caption(issue)
        else:
            st.caption("No data-quality or retrieval issues reported.")
        st.caption(
            "Only successfully retrieved, verified references can be compared. "
            "Similarity percentages do not represent release confidence."
        )
