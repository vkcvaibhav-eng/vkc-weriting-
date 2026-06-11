from __future__ import annotations

import os
from datetime import date

import pandas as pd
import streamlit as st

from logic import (
    AGRO_SANDESH_STYLE_GUIDE,
    DEFAULT_STYLE_SOURCES,
    STYLE_ROLE_LABELS,
    analyze_research_context,
    build_agro_sandesh_docx,
    build_docx,
    cache_default_style_sources,
    citation_key,
    collect_images,
    collect_tables,
    combined_uploaded_text,
    download_and_read_selected_papers,
    extracted_files_from_dict,
    extracted_files_to_dict,
    extract_uploaded_file,
    format_apa_reference,
    generate_agro_sandesh_article,
    generate_full_draft,
    generate_search_queries,
    discover_gujarat_agro_topics,
    load_style_library,
    month_activity_note,
    save_uploaded_style,
    search_and_rank_papers,
    table_to_plain_text,
    truncate_text,
)


st.set_page_config(page_title="Research Paper Writer", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 6px; padding: 0.7rem;}
    .small-muted {color: #64748b; font-size: 0.88rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def init_state() -> None:
    defaults = {
        "openai_key": "",
        "gemini_key": "",
        "gemini_model": "gemini-3.5-flash",
        "agro_month": date.today().strftime("%B"),
        "agro_year": date.today().year,
        "agro_region_scope": "South Gujarat and whole Gujarat",
        "agro_focus_area": "Agricultural acarology and agricultural entomology",
        "agro_crop_filter": "",
        "agro_extra_context": "",
        "agro_local_observations": "",
        "agro_extra_requirements": "",
        "agro_article_length": "1000-1500 words",
        "agro_topic_research": {},
        "agro_article": {},
        "agro_docx_bytes": None,
        "serpapi_key": "",
        "semantic_key": "",
        "core_key": "",
        "model": "gpt-4o-mini",
        "reference_count": 12,
        "per_query_limit": 8,
        "use_ai_scoring": True,
        "extracted_files": [],
        "style_library": {},
        "style_statuses": [],
        "analysis": {},
        "queries": [],
        "paper_search": {},
        "selected_papers": [],
        "downloaded_references": [],
        "draft": {},
        "docx_bytes": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if not st.session_state.style_library:
        try:
            styles, statuses = load_style_library(include_defaults=True)
            st.session_state.style_library = styles
            st.session_state.style_statuses = statuses
        except Exception as exc:
            st.session_state.style_library = {}
            st.session_state.style_statuses = [{"role": "style", "status": f"error: {exc}"}]


def refresh_styles() -> None:
    styles, statuses = load_style_library(include_defaults=True)
    st.session_state.style_library = styles
    st.session_state.style_statuses = statuses


def current_extracted_files():
    return extracted_files_from_dict(st.session_state.get("extracted_files", []))


def current_input_values() -> dict:
    return {
        "paper_title": st.session_state.get("paper_title", "").strip(),
        "authors": st.session_state.get("authors", "").strip(),
        "affiliation": st.session_state.get("affiliation", "").strip(),
        "target_journal": st.session_state.get("target_journal", "").strip(),
        "research_area": st.session_state.get("research_area", "").strip(),
        "master_context": st.session_state.get("master_context", "").strip(),
        "raw_methodology": st.session_state.get("raw_methodology", "").strip(),
    }


def read_secret_or_env(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
        try:
            if name in st.secrets and str(st.secrets[name]).strip():
                return str(st.secrets[name]).strip()
        except Exception:
            continue
    return ""


def docx_safe_name(value: str, fallback: str) -> str:
    name = "".join(ch if ch.isalnum() or ch in {" ", "_", "-"} else "_" for ch in (value or fallback))
    name = "_".join(name.split())
    return (name[:80] or fallback) + ".docx"


def source_list(sources: list[dict]) -> None:
    if not sources:
        st.info("No grounding sources were returned for this step.")
        return
    for index, source in enumerate(sources, start=1):
        title = source.get("title") or f"Source {index}"
        url = source.get("url", "")
        if url:
            st.markdown(f"{index}. [{title}]({url})")
        else:
            st.write(f"{index}. {title}")


def render_agro_sandesh_app() -> None:
    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]
    focus_options = [
        "Agricultural acarology and agricultural entomology",
        "Agricultural entomology",
        "Agricultural acarology",
    ]

    with st.sidebar:
        st.header("Gemini Settings")
        saved_key = st.session_state.gemini_key or read_secret_or_env("GEMINI_API_KEY", "GOOGLE_API_KEY")
        st.session_state.gemini_key = st.text_input("Gemini API key", value=saved_key, type="password")
        st.session_state.gemini_model = st.text_input("Gemini model", value=st.session_state.gemini_model)
        st.divider()
        research = st.session_state.get("agro_topic_research") or {}
        article = st.session_state.get("agro_article") or {}
        st.metric("Topic candidates", len(research.get("topic_candidates", [])))
        st.metric("Research sources", len(research.get("sources", [])))
        st.metric("Article ready", "Yes" if article.get("article_gujarati") else "No")

    st.title("Agro Sandesh Gujarati Article Builder")
    st.caption(
        "Find a timely Gujarat acarology/entomology topic with Gemini + Google Search grounding, "
        "then draft a farmer-centric Gujarati extension article."
    )

    settings, notes = st.columns([0.48, 0.52], gap="large")
    with settings:
        st.subheader("Research Focus")
        col_a, col_b = st.columns(2)
        with col_a:
            st.session_state.agro_month = st.selectbox(
                "Target month",
                months,
                index=months.index(st.session_state.agro_month) if st.session_state.agro_month in months else 0,
            )
        with col_b:
            st.session_state.agro_year = st.number_input(
                "Year",
                min_value=2024,
                max_value=2035,
                value=int(st.session_state.agro_year),
                step=1,
            )
        st.session_state.agro_region_scope = st.text_input(
            "Region scope",
            value=st.session_state.agro_region_scope,
            placeholder="South Gujarat and whole Gujarat",
        )
        st.session_state.agro_focus_area = st.selectbox(
            "Discipline focus",
            focus_options,
            index=focus_options.index(st.session_state.agro_focus_area)
            if st.session_state.agro_focus_area in focus_options
            else 0,
        )
        st.session_state.agro_crop_filter = st.text_input(
            "Crop or commodity preference",
            value=st.session_state.agro_crop_filter,
            placeholder="Optional, e.g., cotton, paddy, mango, vegetables",
        )
        st.info(f"Seasonal clue: {month_activity_note(st.session_state.agro_month)}")

    with notes:
        st.subheader("Local Context")
        st.session_state.agro_extra_context = st.text_area(
            "Extra research context",
            value=st.session_state.agro_extra_context,
            height=150,
            placeholder=(
                "Optional: district, crop stage, farmer complaints, pest/mite observations, "
                "rainfall pattern, market-quality issue, or extension priority."
            ),
        )
        st.session_state.agro_local_observations = st.text_area(
            "Field observations for article",
            value=st.session_state.agro_local_observations,
            height=150,
            placeholder="Optional: field symptoms, farmer questions, local practices, or advisory points to include.",
        )

    if st.button("Deep research and select best topic", type="primary", width="stretch"):
        if not st.session_state.gemini_key:
            st.error("Enter a Gemini API key in the sidebar, or add GEMINI_API_KEY to Streamlit secrets.")
        else:
            with st.spinner("Researching current Gujarat topics with Gemini and Google Search grounding..."):
                try:
                    st.session_state.agro_topic_research = discover_gujarat_agro_topics(
                        api_key=st.session_state.gemini_key,
                        model=st.session_state.gemini_model,
                        month=st.session_state.agro_month,
                        year=int(st.session_state.agro_year),
                        region_scope=st.session_state.agro_region_scope,
                        focus_area=st.session_state.agro_focus_area,
                        crop_filter=st.session_state.agro_crop_filter,
                        extra_context=st.session_state.agro_extra_context,
                    )
                    st.session_state.agro_article = {}
                    st.session_state.agro_docx_bytes = None
                    st.success("Topic research completed.")
                except Exception as exc:
                    st.error(str(exc))

    research = st.session_state.get("agro_topic_research") or {}
    if research:
        selected = research.get("selected_topic") or {}
        st.divider()
        st.subheader("Selected Topic")
        score_cols = st.columns(4)
        score_cols[0].metric("Score", selected.get("score", ""))
        score_cols[1].metric("Focus", selected.get("focus_area", st.session_state.agro_focus_area))
        score_cols[2].metric("Crop", selected.get("crop", ""))
        score_cols[3].metric("Research date", research.get("research_date", ""))
        st.markdown(f"**{selected.get('topic', 'Selected topic')}**")
        if selected.get("farmer_problem"):
            st.write(selected["farmer_problem"])
        if research.get("selection_reason"):
            st.caption(research["selection_reason"])

        candidates = research.get("topic_candidates") or []
        if candidates:
            candidate_rows = [
                {
                    "score": item.get("score", ""),
                    "topic": item.get("topic", ""),
                    "crop": item.get("crop", ""),
                    "focus": item.get("focus_area", ""),
                    "month_relevance": item.get("month_relevance", ""),
                }
                for item in candidates
            ]
            with st.expander("Candidate topic ranking", expanded=True):
                st.dataframe(pd.DataFrame(candidate_rows), width="stretch", hide_index=True)

        with st.expander("Grounded research sources", expanded=False):
            source_list(research.get("sources", []))
            queries = research.get("search_queries", [])
            if queries:
                st.write("Search queries used")
                for query in queries:
                    st.code(query, language="text")

        st.divider()
        st.subheader("Article Draft")
        controls = st.columns([0.32, 0.68], gap="large")
        with controls[0]:
            st.session_state.agro_article_length = st.selectbox(
                "Target length",
                ["800-1000 words", "1000-1500 words", "1500-1800 words"],
                index=["800-1000 words", "1000-1500 words", "1500-1800 words"].index(
                    st.session_state.agro_article_length
                )
                if st.session_state.agro_article_length in ["800-1000 words", "1000-1500 words", "1500-1800 words"]
                else 1,
            )
            with st.expander("Writing rules", expanded=False):
                st.text(AGRO_SANDESH_STYLE_GUIDE)
        with controls[1]:
            st.session_state.agro_extra_requirements = st.text_area(
                "Extra article requirements",
                value=st.session_state.agro_extra_requirements,
                height=130,
                placeholder="Optional: points to emphasize, avoid, or include for Agro Sandesh.",
            )

        if st.button("Generate Gujarati Agro Sandesh article", type="primary", width="stretch"):
            if not st.session_state.gemini_key:
                st.error("Enter a Gemini API key in the sidebar, or add GEMINI_API_KEY to Streamlit secrets.")
            else:
                with st.spinner("Writing the Gujarati farmer-extension article..."):
                    try:
                        st.session_state.agro_article = generate_agro_sandesh_article(
                            api_key=st.session_state.gemini_key,
                            model=st.session_state.gemini_model,
                            topic_context=research,
                            month=st.session_state.agro_month,
                            year=int(st.session_state.agro_year),
                            region_scope=st.session_state.agro_region_scope,
                            focus_area=st.session_state.agro_focus_area,
                            article_length=st.session_state.agro_article_length,
                            local_observations=st.session_state.agro_local_observations,
                            extra_requirements=st.session_state.agro_extra_requirements,
                        )
                        st.session_state.agro_docx_bytes = None
                        st.success("Article generated.")
                    except Exception as exc:
                        st.error(str(exc))

    article = st.session_state.get("agro_article") or {}
    if article.get("article_gujarati"):
        st.divider()
        st.markdown(f"### {article.get('title_gujarati', 'Agro Sandesh Article')}")
        st.markdown(article["article_gujarati"])

        if article.get("practical_takeaway_gujarati"):
            st.success(article["practical_takeaway_gujarati"])

        checklist = article.get("final_checklist") or {}
        if checklist:
            with st.expander("Final writing checklist", expanded=False):
                st.dataframe(
                    pd.DataFrame(
                        [{"check": key.replace("_", " "), "passed": bool(value)} for key, value in checklist.items()]
                    ),
                    width="stretch",
                    hide_index=True,
                )

        with st.expander("Article sources", expanded=False):
            source_list(article.get("sources", []))
            if article.get("source_use_note"):
                st.write(article["source_use_note"])

        if st.button("Build article DOCX", width="stretch"):
            try:
                st.session_state.agro_docx_bytes = build_agro_sandesh_docx(article)
                st.success("DOCX is ready.")
            except Exception as exc:
                st.error(str(exc))

        if st.session_state.get("agro_docx_bytes"):
            st.download_button(
                "Download Agro Sandesh DOCX",
                data=st.session_state.agro_docx_bytes,
                file_name=docx_safe_name(article.get("title_gujarati", ""), "agro_sandesh_article"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )


def build_context_for_search(inputs: dict, files) -> tuple[dict, str, list[str]]:
    result_text = combined_uploaded_text(files)
    analysis = analyze_research_context(
        st.session_state.openai_key,
        st.session_state.model,
        inputs["paper_title"],
        inputs["research_area"],
        inputs["master_context"],
        inputs["raw_methodology"],
        result_text,
    )
    queries = generate_search_queries(
        st.session_state.openai_key,
        st.session_state.model,
        analysis,
        inputs["master_context"] + "\n" + inputs["raw_methodology"],
        result_text,
        max_queries=6,
    )
    return analysis, result_text, queries


def paper_rows(papers: list[dict]) -> pd.DataFrame:
    rows = []
    for index, paper in enumerate(papers):
        rows.append(
            {
                "selected": bool(paper.get("selected")),
                "rank": index + 1,
                "score": paper.get("score", 0),
                "category": paper.get("category", "Research Article"),
                "citation": citation_key(paper),
                "year": paper.get("year"),
                "citations": paper.get("citation_count", 0),
                "pdf": "yes" if paper.get("pdf_url") or paper.get("pdf_urls") else "",
                "full_text_chars": paper.get("full_text_chars", 0),
                "source": paper.get("source", ""),
                "title": paper.get("title", ""),
                "reason": paper.get("score_reason", ""),
                "paper_id": paper.get("paper_id", ""),
            }
        )
    return pd.DataFrame(rows)


init_state()

app_mode = st.sidebar.radio(
    "Workflow",
    ["Agro Sandesh Gujarati Article", "Research Paper Writer"],
    key="app_mode",
)

if app_mode == "Agro Sandesh Gujarati Article":
    render_agro_sandesh_app()
    st.stop()

with st.sidebar:
    st.header("API Settings")
    openai_default = st.session_state.openai_key or read_secret_or_env("OPENAI_API_KEY", "OpenAI_key")
    openai_key = st.text_input("OpenAI API key", type="password", value=openai_default)
    if openai_key != st.session_state.openai_key:
        st.session_state.openai_key = openai_key

    st.session_state.model = st.text_input("OpenAI model", value=st.session_state.model)

    st.divider()
    st.subheader("Reference Search")
    serpapi_key = st.text_input(
        "SerpAPI key",
        type="password",
        value=st.session_state.serpapi_key or read_secret_or_env("SERPAPI_API_KEY", "SERPAPI_KEY"),
    )
    semantic_key = st.text_input(
        "Semantic Scholar key",
        type="password",
        value=st.session_state.semantic_key or read_secret_or_env("SEMANTIC_SCHOLAR_API_KEY"),
    )
    core_key = st.text_input(
        "CORE API key",
        type="password",
        value=st.session_state.core_key or read_secret_or_env("CORE_API_KEY"),
    )
    st.session_state.serpapi_key = serpapi_key
    st.session_state.semantic_key = semantic_key
    st.session_state.core_key = core_key

    st.session_state.reference_count = st.slider("References to select", min_value=3, max_value=25, value=st.session_state.reference_count)
    st.session_state.per_query_limit = st.slider("Results per query", min_value=3, max_value=15, value=st.session_state.per_query_limit)
    st.session_state.use_ai_scoring = st.checkbox("Use AI scoring", value=st.session_state.use_ai_scoring)

    st.divider()
    loaded_roles = sum(1 for text in st.session_state.style_library.values() if text)
    st.metric("Loaded style roles", f"{loaded_roles}/{len(STYLE_ROLE_LABELS)}")
    extracted_count = len(st.session_state.get("extracted_files", []))
    st.metric("Extracted input files", extracted_count)
    downloaded_count = len([p for p in st.session_state.get("downloaded_references", []) if p.get("download_success")])
    st.metric("Full-text references", downloaded_count)


st.title("Research Paper Writing App")
st.caption("One-page workflow for methodology rewriting, result writing, discussion evidence, APA 7 references, and DOCX export.")

tabs = st.tabs(["Inputs", "Style Library", "Evidence Search", "Draft Preview", "Export"])


with tabs[0]:
    left, right = st.columns([1.05, 0.95], gap="large")
    with left:
        st.subheader("Paper Information")
        st.text_input("Paper title", key="paper_title", placeholder="Leave blank to auto-generate from methodology and results")
        st.text_input("Scientist / author names", key="authors", placeholder="e.g., A. B. Patel, C. D. Parmar")
        st.text_input("Affiliation / notes", key="affiliation", placeholder="Department, college, university")
        meta_cols = st.columns(2)
        with meta_cols[0]:
            st.text_input("Target journal", key="target_journal")
        with meta_cols[1]:
            st.text_input("Research area / discipline", key="research_area", placeholder="Agricultural entomology")

        st.subheader("Core Inputs")
        st.text_area(
            "Master context",
            key="master_context",
            height=180,
            placeholder=(
                "Crop / host, pest or pathogen, location, season, design, treatments, replications, "
                "observations, key factual outcomes."
            ),
        )
        st.text_area(
            "Raw methodology",
            key="raw_methodology",
            height=260,
            placeholder="Paste the unpolished methodology/protocol here.",
        )

    with right:
        st.subheader("Result Files")
        result_files = st.file_uploader(
            "Upload result tables/docs",
            type=["docx", "pdf", "csv", "xlsx", "xls", "txt", "md"],
            accept_multiple_files=True,
        )
        graph_files = st.file_uploader(
            "Upload graph/image files",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
        )

        if st.button("Extract uploaded files", type="primary", width="stretch"):
            files_to_process = list(result_files or []) + list(graph_files or [])
            if not files_to_process:
                st.warning("Upload at least one result, table, document, or graph file.")
            else:
                extracted = []
                errors = []
                with st.spinner("Extracting text, tables, and graph summaries..."):
                    for uploaded in files_to_process:
                        try:
                            extracted.append(
                                extract_uploaded_file(uploaded, st.session_state.openai_key, st.session_state.model)
                            )
                        except Exception as exc:
                            errors.append(f"{uploaded.name}: {exc}")
                st.session_state.extracted_files = extracted_files_to_dict(extracted)
                st.session_state.docx_bytes = None
                st.session_state.draft = {}
                if errors:
                    st.error("\n".join(errors))
                st.success(f"Extracted {len(extracted)} file(s).")

        extracted_files = current_extracted_files()
        if extracted_files:
            st.subheader("Extraction Summary")
            summary_rows = []
            for item in extracted_files:
                summary_rows.append(
                    {
                        "file": item.name,
                        "type": item.kind,
                        "characters": len(item.text or ""),
                        "tables": len(item.tables or []),
                        "image": "yes" if item.image_bytes else "",
                    }
                )
            st.dataframe(pd.DataFrame(summary_rows), width="stretch", hide_index=True)

            tables = collect_tables(extracted_files)
            if tables:
                with st.expander("Table preview", expanded=False):
                    st.text(truncate_text("\n\n".join(table_to_plain_text(t) for t in tables[:3]), 5000))

            images = collect_images(extracted_files)
            if images:
                with st.expander("Graph/image preview", expanded=False):
                    for image in images:
                        st.image(image["bytes"], caption=image["name"], width="stretch")
                        if image.get("summary"):
                            st.write(image["summary"])


with tabs[1]:
    st.subheader("Permanent Style Library")
    col_a, col_b = st.columns([0.58, 0.42], gap="large")

    with col_a:
        if st.button("Refresh default PDF styles", width="stretch"):
            with st.spinner("Reading default style PDFs..."):
                refresh_styles()
            st.success("Default style library refreshed.")

        status_rows = []
        for status in st.session_state.style_statuses:
            status_rows.append(
                {
                    "role": STYLE_ROLE_LABELS.get(status.get("role"), status.get("role")),
                    "status": status.get("status"),
                    "characters": status.get("characters", 0),
                    "source": status.get("source", ""),
                }
            )
        if status_rows:
            st.dataframe(pd.DataFrame(status_rows), width="stretch", hide_index=True)

        with st.expander("Default source paths", expanded=False):
            for role, meta in DEFAULT_STYLE_SOURCES.items():
                st.markdown(f"**{STYLE_ROLE_LABELS.get(role, role)}**")
                st.code(meta["path"], language="text")
                st.caption(meta["role"])

    with col_b:
        st.write("Add style example")
        upload_role = st.selectbox(
            "Style role",
            options=list(STYLE_ROLE_LABELS.keys()),
            format_func=lambda value: STYLE_ROLE_LABELS[value],
        )
        style_file = st.file_uploader("Upload PDF/DOCX/TXT style file", type=["pdf", "docx", "txt", "md"])
        if st.button("Save style example", width="stretch"):
            if not style_file:
                st.warning("Choose a style file first.")
            else:
                try:
                    result = save_uploaded_style(style_file, upload_role)
                    refresh_styles()
                    st.success(f"Saved {result['characters']} characters for {STYLE_ROLE_LABELS[upload_role]}.")
                except Exception as exc:
                    st.error(str(exc))

        st.write("Style preview")
        preview_role = st.selectbox(
            "Preview role",
            options=list(STYLE_ROLE_LABELS.keys()),
            format_func=lambda value: STYLE_ROLE_LABELS[value],
            key="preview_role",
        )
        preview_text = st.session_state.style_library.get(preview_role, "")
        st.text_area("Loaded text", value=truncate_text(preview_text, 3000), height=260, disabled=True)


with tabs[2]:
    st.subheader("Reference Search, Sorting, Filtering, and PDF Reading")
    st.caption(
        "This combines the old search engine, sorting/filtering dashboard, and PDF downloader inside the writing app."
    )
    inputs = current_input_values()
    extracted_files = current_extracted_files()

    top = st.columns([0.2, 0.2, 0.2, 0.2, 0.2])
    top[0].metric("Reference target", st.session_state.reference_count)
    top[1].metric("Result files", len(extracted_files))
    top[2].metric("Selected", len(st.session_state.get("selected_papers", [])))
    top[3].metric(
        "PDFs read",
        len([p for p in st.session_state.get("downloaded_references", []) if p.get("download_success")]),
    )
    top[4].metric("AI scoring", "On" if st.session_state.use_ai_scoring else "Off")

    if st.button("Analyze findings and search papers", type="primary", width="stretch"):
        if not (inputs["master_context"] or inputs["raw_methodology"] or extracted_files):
            st.warning("Add methodology, master context, or result files before searching.")
        else:
            with st.spinner("Analyzing findings, building queries, searching, deduplicating, and scoring papers..."):
                analysis, result_text, queries = build_context_for_search(inputs, extracted_files)
                context_text = "\n".join(
                    [
                        inputs["paper_title"],
                        inputs["research_area"],
                        inputs["master_context"],
                        inputs["raw_methodology"],
                        result_text,
                        " ".join(map(str, analysis.get("major_findings", []))),
                    ]
                )
                search_result = search_and_rank_papers(
                    queries=queries,
                    context_text=context_text,
                    semantic_key=st.session_state.semantic_key,
                    serpapi_key=st.session_state.serpapi_key,
                    core_key=st.session_state.core_key,
                    openai_key=st.session_state.openai_key,
                    model=st.session_state.model,
                    reference_count=st.session_state.reference_count,
                    per_query_limit=st.session_state.per_query_limit,
                    use_ai_scoring=st.session_state.use_ai_scoring,
                )
            st.session_state.analysis = analysis
            st.session_state.queries = queries
            st.session_state.paper_search = search_result
            st.session_state.selected_papers = search_result.get("selected", [])
            st.session_state.downloaded_references = []
            st.success(
                f"Found {search_result.get('candidate_count', 0)} candidates, "
                f"{search_result.get('deduped_count', 0)} after deduplication."
            )

    if st.session_state.get("queries"):
        with st.expander("Search queries", expanded=True):
            for query in st.session_state.queries:
                st.code(query, language="text")

    search_result = st.session_state.get("paper_search") or {}
    warnings = search_result.get("warnings", [])
    if warnings:
        with st.expander("Search warnings", expanded=False):
            for warning in warnings:
                st.warning(warning)

    papers = search_result.get("papers", [])
    if papers:
        df = paper_rows(papers)
        category_counts = df["category"].value_counts().to_dict() if "category" in df else {}
        count_cols = st.columns(3)
        count_cols[0].metric("Research articles", category_counts.get("Research Article", 0))
        count_cols[1].metric("Review papers", category_counts.get("Review Paper", 0))
        count_cols[2].metric("Theses", category_counts.get("Thesis", 0))
        st.caption("Sort columns in the table and tick Use for references that should support the paper.")
        edited = st.data_editor(
            df,
            width="stretch",
            hide_index=True,
            disabled=[
                "rank",
                "score",
                "category",
                "citation",
                "year",
                "citations",
                "pdf",
                "full_text_chars",
                "source",
                "title",
                "reason",
                "paper_id",
            ],
            column_config={"selected": st.column_config.CheckboxColumn("Use")},
        )
        if st.button("Save selected papers", width="stretch"):
            selected_ids = set(edited.loc[edited["selected"] == True, "paper_id"].astype(str).tolist())
            selected = []
            for paper in papers:
                paper["selected"] = str(paper.get("paper_id")) in selected_ids
                if paper["selected"]:
                    selected.append(paper)
            st.session_state.paper_search["papers"] = papers
            st.session_state.selected_papers = selected
            st.session_state.downloaded_references = [
                item for item in st.session_state.get("downloaded_references", []) if str(item.get("paper_id")) in selected_ids
            ]
            st.success(f"Saved {len(selected)} selected paper(s).")

        selected_papers = st.session_state.get("selected_papers", [])
        if selected_papers:
            st.divider()
            st.subheader("Download and Read Selected Papers")
            st.caption(
                "Use this before drafting so the discussion can use full-paper text, not only title or abstract."
            )
            dl_cols = st.columns(4)
            dl_cols[0].metric("Selected", len(selected_papers))
            dl_cols[1].metric("With PDF clue", len([p for p in selected_papers if p.get("pdf_url") or p.get("pdf_urls")]))
            downloaded_refs = st.session_state.get("downloaded_references", [])
            dl_cols[2].metric("Downloaded", len([p for p in downloaded_refs if p.get("download_success")]))
            dl_cols[3].metric("Readable chars", sum(int(p.get("full_text_chars") or 0) for p in downloaded_refs))

            if st.button("Download and read selected PDFs", type="primary", width="stretch"):
                with st.spinner("Finding open PDFs, downloading selected papers, and extracting full text..."):
                    downloaded = download_and_read_selected_papers(
                        selected_papers,
                        serpapi_key=st.session_state.serpapi_key,
                        core_key=st.session_state.core_key,
                        semantic_key=st.session_state.semantic_key,
                    )
                by_id = {str(item.get("paper_id")): item for item in downloaded}
                st.session_state.selected_papers = [by_id.get(str(p.get("paper_id")), p) for p in selected_papers]
                st.session_state.downloaded_references = downloaded
                if st.session_state.paper_search.get("papers"):
                    updated_papers = []
                    for paper in st.session_state.paper_search["papers"]:
                        updated_papers.append(by_id.get(str(paper.get("paper_id")), paper))
                    st.session_state.paper_search["papers"] = updated_papers
                success_count = len([p for p in downloaded if p.get("download_success")])
                st.success(f"Downloaded/read {success_count}/{len(downloaded)} selected paper(s).")

            downloaded_refs = st.session_state.get("downloaded_references", [])
            if downloaded_refs:
                download_rows = [
                    {
                        "status": "downloaded" if item.get("download_success") else "not found",
                        "method": item.get("download_method", ""),
                        "chars": item.get("full_text_chars", 0),
                        "size_mb": round((item.get("pdf_size_bytes", 0) or 0) / 1024 / 1024, 2),
                        "category": item.get("category", ""),
                        "title": item.get("title", ""),
                        "pdf_path": item.get("pdf_path", ""),
                    }
                    for item in downloaded_refs
                ]
                st.dataframe(pd.DataFrame(download_rows), width="stretch", hide_index=True)

            with st.expander("APA 7 reference preview", expanded=False):
                for paper in st.session_state.get("selected_papers", []):
                    st.write(format_apa_reference(paper))


with tabs[3]:
    st.subheader("Draft Preview")
    inputs = current_input_values()
    extracted_files = current_extracted_files()
    selected_papers = st.session_state.get("selected_papers", [])

    cols = st.columns(4)
    cols[0].metric("Selected references", len(selected_papers))
    cols[1].metric("Full-text references", len([p for p in selected_papers if p.get("download_success")]))
    cols[2].metric("Tables", len(collect_tables(extracted_files)))
    cols[3].metric("Graphs", len(collect_images(extracted_files)))

    if st.button("Generate full paper draft", type="primary", width="stretch"):
        if not st.session_state.openai_key:
            st.error("Enter an OpenAI API key in the sidebar before drafting.")
        elif not (inputs["master_context"] or inputs["raw_methodology"] or extracted_files):
            st.warning("Add methodology, master context, or result files before drafting.")
        else:
            if selected_papers and not any(p.get("download_success") for p in selected_papers):
                st.warning("Selected references have not been downloaded/read yet; discussion will rely mostly on abstracts.")
            with st.spinner("Writing methodology, results, discussion, conclusion, introduction, and abstract..."):
                try:
                    draft = generate_full_draft(
                        api_key=st.session_state.openai_key,
                        model=st.session_state.model,
                        paper_title=inputs["paper_title"],
                        authors=inputs["authors"],
                        affiliation=inputs["affiliation"],
                        research_area=inputs["research_area"],
                        master_context=inputs["master_context"],
                        raw_methodology=inputs["raw_methodology"],
                        extracted_files=extracted_files,
                        styles=st.session_state.style_library,
                        selected_papers=selected_papers,
                    )
                    st.session_state.draft = draft
                    st.session_state.docx_bytes = None
                    st.success("Draft generated.")
                except Exception as exc:
                    st.error(str(exc))

    draft = st.session_state.get("draft") or {}
    if draft:
        st.markdown(f"### {draft.get('title', 'Research Paper Draft')}")
        if draft.get("authors"):
            st.caption(draft["authors"])
        section_map = [
            ("Abstract", "abstract"),
            ("Introduction", "introduction"),
            ("Materials and Methods", "methodology"),
            ("Results", "results"),
            ("Discussion", "discussion"),
            ("Conclusion", "conclusion"),
            ("References", "references"),
        ]
        for label, key in section_map:
            with st.expander(label, expanded=label in {"Abstract", "Results"}):
                value = draft.get(key, "")
                if key == "references":
                    for ref in value:
                        st.write(ref)
                else:
                    st.write(value)


with tabs[4]:
    st.subheader("Export")
    draft = st.session_state.get("draft") or {}
    if not draft:
        st.info("Generate a draft before exporting.")
    else:
        export_name = st.text_input(
            "DOCX file name",
            value=(draft.get("title") or "research_paper").replace(" ", "_")[:80] + ".docx",
        )
        if st.button("Build DOCX", type="primary", width="stretch"):
            try:
                with st.spinner("Building Word document..."):
                    st.session_state.docx_bytes = build_docx(draft)
                st.success("DOCX is ready.")
            except Exception as exc:
                st.error(str(exc))

        if st.session_state.docx_bytes:
            st.download_button(
                "Download research paper DOCX",
                data=st.session_state.docx_bytes,
                file_name=export_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

        st.divider()
        st.write("Export contents")
        export_rows = [
            {"section": "Title", "characters": len(draft.get("title", ""))},
            {"section": "Abstract", "characters": len(draft.get("abstract", ""))},
            {"section": "Introduction", "characters": len(draft.get("introduction", ""))},
            {"section": "Materials and Methods", "characters": len(draft.get("methodology", ""))},
            {"section": "Results", "characters": len(draft.get("results", ""))},
            {"section": "Discussion", "characters": len(draft.get("discussion", ""))},
            {"section": "Conclusion", "characters": len(draft.get("conclusion", ""))},
            {"section": "References", "characters": sum(len(ref) for ref in draft.get("references", []))},
        ]
        st.dataframe(pd.DataFrame(export_rows), width="stretch", hide_index=True)
