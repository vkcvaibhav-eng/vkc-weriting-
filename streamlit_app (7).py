from __future__ import annotations

import importlib
import os

import pandas as pd
import streamlit as st

import logic as logic_module

importlib.reload(logic_module)

from logic import (
    DISCUSSION_WORKFLOW_TEXT,
    DEFAULT_CLAUDE_MODEL,
    DEFAULT_GEMINI_MODEL,
    DEFAULT_PERPLEXITY_MODEL,
    DEFAULT_SHELTON_STYLE_PATH,
    MIN_REFERENCE_COUNT,
    STYLE_PRESETS,
    audit_draft_with_gemini_style_editor,
    audit_draft_with_openai_style_editor,
    analyze_research_context,
    build_discussion_framework,
    build_docx,
    citation_key,
    collect_images,
    collect_tables,
    combined_uploaded_text,
    download_and_read_selected_papers,
    extracted_files_from_dict,
    extracted_files_to_dict,
    extract_uploaded_file,
    format_apa_reference,
    gemini_read_selected_papers,
    generate_full_draft,
    generate_claude_discussion_search_plan,
    generate_search_queries,
    include_in_final_references,
    load_sau_icar_results_prompt,
    load_writing_style_contract,
    merge_search_results,
    recommend_writing_style,
    revise_draft_with_style_audits,
    review_primary_study_search_queries,
    search_and_rank_papers,
    section_prompt_common,
    suggest_style_aligned_followup_needs,
    summarize_gemini_reference_notes,
    table_to_plain_text,
    thesis_primary_study_search_queries,
    truncate_text,
    writing_length_directive,
    write_results,
)


st.set_page_config(page_title="Research Paper Writer", layout="wide")


API_SECRET_NAMES = {
    "openai_key": (
        "OPENAI_API_KEY",
        "OPENAI_KEY",
        "openai_api_key",
        "openai.key",
        "openai.api_key",
    ),
    "serpapi_key": (
        "SERPAPI_API_KEY",
        "SERPAPI_KEY",
        "serpapi_api_key",
        "serpapi.key",
        "serpapi.api_key",
    ),
    "semantic_key": (
        "SEMANTIC_SCHOLAR_API_KEY",
        "SEMANTIC_KEY",
        "semantic_scholar_api_key",
        "semantic_scholar.key",
        "semantic_scholar.api_key",
    ),
    "core_key": (
        "CORE_API_KEY",
        "CORE_KEY",
        "core_api_key",
        "core.key",
        "core.api_key",
    ),
    "perplexity_key": (
        "PERPLEXITY_API_KEY",
        "PERPLEXITY_KEY",
        "perplexity_api_key",
        "perplexity.key",
        "perplexity.api_key",
    ),
    "gemini_key": (
        "GEMINI_API_KEY",
        "GOOGLE_API_KEY",
        "GEMINI_KEY",
        "gemini_api_key",
        "google_api_key",
        "gemini.key",
        "gemini.api_key",
    ),
    "claude_key": (
        "ANTHROPIC_API_KEY",
        "CLAUDE_API_KEY",
        "ANTHROPIC_KEY",
        "CLAUDE_KEY",
        "anthropic_api_key",
        "claude_api_key",
        "anthropic.key",
        "anthropic.api_key",
        "claude.key",
        "claude.api_key",
    ),
}

API_SECRETS_TEMPLATE = """OPENAI_API_KEY = "your-openai-key"
SERPAPI_API_KEY = "your-serpapi-key"
SEMANTIC_SCHOLAR_API_KEY = "your-semantic-scholar-key"
CORE_API_KEY = "your-core-key"
PERPLEXITY_API_KEY = "your-perplexity-key"
GEMINI_API_KEY = "your-gemini-key"
ANTHROPIC_API_KEY = "your-claude-key"
"""


def read_streamlit_secret(name: str) -> str:
    try:
        current = st.secrets
        for part in name.split("."):
            current = current[part]
        return str(current).strip() if current else ""
    except Exception:
        return ""


def read_secret_or_env(*names: str) -> str:
    for name in names:
        value = read_streamlit_secret(name)
        if value:
            return value
    for name in names:
        env_name = name.upper().replace(".", "_")
        value = os.getenv(env_name, "")
        if value:
            return value.strip()
    return ""


def sync_api_keys_from_secrets() -> None:
    if st.session_state.get("manual_api_key_override"):
        return
    for state_key, secret_names in API_SECRET_NAMES.items():
        value = read_secret_or_env(*secret_names)
        if value:
            st.session_state[state_key] = value


def configure_api_key(label: str, state_key: str) -> None:
    secret_value = read_secret_or_env(*API_SECRET_NAMES[state_key])
    if secret_value and not st.session_state.get("manual_api_key_override"):
        st.session_state[state_key] = secret_value
        st.caption(f"{label}: loaded from Streamlit secrets")
        return

    value = st.text_input(label, type="password", value="" if secret_value else st.session_state[state_key])
    if value:
        st.session_state[state_key] = value
    elif not secret_value:
        st.session_state[state_key] = ""


def init_state() -> None:
    defaults = {
        "openai_key": "",
        "serpapi_key": "",
        "semantic_key": "",
        "core_key": "",
        "perplexity_key": "",
        "perplexity_model": DEFAULT_PERPLEXITY_MODEL,
        "gemini_key": "",
        "gemini_model": DEFAULT_GEMINI_MODEL,
        "claude_key": "",
        "claude_model": DEFAULT_CLAUDE_MODEL,
        "model": "gpt-4o-mini",
        "manual_api_key_override": False,
        "reference_count": MIN_REFERENCE_COUNT,
        "per_query_limit": 10,
        "use_ai_scoring": True,
        "writing_length_mode": "Concise style-preserving",
        "extracted_files": [],
        "active_style_name": "Anthony M. Shelton",
        "active_style_path": DEFAULT_SHELTON_STYLE_PATH,
        "active_style_contract": {},
        "active_style_report": {},
        "style_recommendation": {},
        "analysis": {},
        "queries": [],
        "paper_search": {},
        "selected_papers": [],
        "downloaded_references": [],
        "gemini_recommendations": {},
        "followup_suggestions": {},
        "openai_style_audit": {},
        "gemini_style_audit": {},
        "draft": {},
        "docx_bytes": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.reference_count < MIN_REFERENCE_COUNT:
        st.session_state.reference_count = MIN_REFERENCE_COUNT
    if st.session_state.get("shelton_style_contract") and not st.session_state.get("active_style_contract"):
        st.session_state.active_style_contract = st.session_state.shelton_style_contract
        st.session_state.active_style_report = st.session_state.get("shelton_style_report", {})
    sync_api_keys_from_secrets()


def current_extracted_files():
    return extracted_files_from_dict(st.session_state.get("extracted_files", []))


def current_style_contract() -> dict:
    return st.session_state.get("active_style_contract") or {}


def current_style_profile() -> dict:
    return (current_style_contract().get("profile") or {})


def current_style_library() -> dict:
    return (current_style_contract().get("styles") or {})


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


def build_context_for_search(inputs: dict, files) -> tuple[dict, str, list[str]]:
    result_text = combined_uploaded_text(files)
    context_text = inputs["master_context"] + "\n" + inputs["raw_methodology"]
    analysis = analyze_research_context(
        st.session_state.openai_key,
        st.session_state.model,
        inputs["paper_title"],
        inputs["research_area"],
        inputs["master_context"],
        inputs["raw_methodology"],
        result_text,
    )
    result_title = inputs["paper_title"] or analysis.get("generated_title") or "Research Paper Draft"
    analysis["analysis_title"] = result_title
    styles = current_style_library()
    common = section_prompt_common(
        result_title,
        inputs.get("authors", ""),
        inputs.get("affiliation", ""),
        inputs.get("research_area", ""),
        inputs.get("master_context", ""),
        analysis,
        result_text,
        writing_length_directive(st.session_state.writing_length_mode),
    )
    tables = collect_tables(files)
    images = collect_images(files)
    table_summaries = "\n\n".join(table_to_plain_text(table) for table in tables)
    image_summaries = "\n\n".join(
        f"Figure from {image['name']}:\n{image.get('summary', '')}" for image in images if image.get("summary")
    )

    results_section = ""
    if st.session_state.openai_key:
        try:
            results_section = write_results(
                st.session_state.openai_key,
                st.session_state.model,
                common,
                styles,
                table_summaries,
                image_summaries,
            )
            analysis["results_draft"] = results_section
        except Exception as exc:
            analysis["results_draft_error"] = str(exc)
    else:
        analysis["results_draft_error"] = "OpenAI key is required to write the first Results draft."

    discussion_framework = build_discussion_framework(
        st.session_state.openai_key,
        st.session_state.model,
        common,
        styles,
        analysis,
        [],
        results_section or result_text,
        claude_key=st.session_state.claude_key,
        claude_model=st.session_state.claude_model,
    )
    analysis["discussion_framework"] = discussion_framework

    base_queries = generate_search_queries(
        st.session_state.openai_key,
        st.session_state.model,
        analysis,
        context_text + "\n\nDrafted Results:\n" + results_section,
        result_text,
        max_queries=6,
    )
    claude_plan = generate_claude_discussion_search_plan(
        st.session_state.claude_key,
        st.session_state.claude_model,
        analysis,
        context_text,
        result_text,
        current_style_profile(),
        results_section=results_section,
        discussion_framework=discussion_framework,
        max_queries=8,
    )
    claude_queries = claude_plan.get("discussion_search_queries") or []
    if claude_plan.get("needed_paper_types") or claude_queries or claude_plan.get("claude_query_error"):
        analysis["claude_discussion_search_plan"] = claude_plan

    queries = []
    seen_queries = set()
    for query_group in (base_queries, claude_queries):
        for query in query_group:
            clean_query = clean_query_text(query)
            normalized_query = " ".join(clean_query.lower().split())
            if clean_query and normalized_query not in seen_queries:
                queries.append(clean_query)
                seen_queries.add(normalized_query)
    return analysis, result_text, queries


def current_context_text(inputs: dict, files, analysis: dict | None = None) -> str:
    result_text = combined_uploaded_text(files)
    analysis = analysis or st.session_state.get("analysis") or {}
    return "\n".join(
        [
            inputs.get("paper_title", ""),
            inputs.get("research_area", ""),
            inputs.get("master_context", ""),
            inputs.get("raw_methodology", ""),
            result_text,
            " ".join(map(str, analysis.get("major_findings", []))),
            " ".join(map(str, analysis.get("keywords", []))),
        ]
    )


def clean_query_text(query) -> str:
    if isinstance(query, dict):
        query = query.get("query") or query.get("search_query") or query.get("title") or ""
    return str(query).strip()


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


PAPER_TABLE_DISABLED_COLUMNS = [
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
]


def paper_selection_editor(papers: list[dict], key: str) -> pd.DataFrame:
    return st.data_editor(
        paper_rows(papers),
        width="stretch",
        hide_index=True,
        disabled=PAPER_TABLE_DISABLED_COLUMNS,
        column_config={"selected": st.column_config.CheckboxColumn("Use")},
        key=key,
    )


init_state()

with st.sidebar:
    st.header("API Settings")
    st.checkbox("Manual API key override", key="manual_api_key_override")
    configure_api_key("OpenAI API key", "openai_key")

    st.session_state.model = st.text_input("OpenAI model", value=st.session_state.model)

    st.divider()
    st.subheader("Reference Search")
    configure_api_key("SerpAPI key", "serpapi_key")
    configure_api_key("Semantic Scholar key", "semantic_key")
    configure_api_key("CORE API key", "core_key")
    configure_api_key("Perplexity Sonar key", "perplexity_key")
    st.session_state.perplexity_model = st.text_input("Perplexity model", value=st.session_state.perplexity_model)
    configure_api_key("Google Gemini API key", "gemini_key")
    st.session_state.gemini_model = st.text_input("Gemini model", value=st.session_state.gemini_model)

    st.divider()
    st.subheader("Premium Discussion")
    configure_api_key("Claude / Anthropic API key", "claude_key")
    st.session_state.claude_model = st.text_input("Claude discussion model", value=st.session_state.claude_model)
    st.caption("Claude is used for the Discussion framework and Discussion section when this key is available.")

    st.divider()
    st.subheader("Writing Length")
    length_modes = ["Concise style-preserving", "Very concise style-preserving", "Standard detailed"]
    current_length_mode = st.session_state.get("writing_length_mode", "Concise style-preserving")
    if current_length_mode not in length_modes:
        current_length_mode = "Concise style-preserving"
    st.session_state.writing_length_mode = st.selectbox(
        "Draft length mode",
        length_modes,
        index=length_modes.index(current_length_mode),
    )
    st.caption("Concise modes keep the loaded author style and analysis, but remove repetition and padding.")

    with st.expander("Streamlit secrets template"):
        st.code(API_SECRETS_TEMPLATE, language="toml")

    st.session_state.reference_count = st.slider(
        "References to select",
        min_value=MIN_REFERENCE_COUNT,
        max_value=30,
        value=max(st.session_state.reference_count, MIN_REFERENCE_COUNT),
    )
    st.session_state.per_query_limit = st.slider("Results per query", min_value=5, max_value=20, value=st.session_state.per_query_limit)
    st.session_state.use_ai_scoring = st.checkbox("Use AI scoring", value=st.session_state.use_ai_scoring)

    st.divider()
    st.metric("Minimum evidence set", "10 papers + 3 theses + 1 review")
    style_status = current_style_profile().get("author") if current_style_contract() else "Not loaded"
    st.metric("Writing style", style_status)
    extracted_count = len(st.session_state.get("extracted_files", []))
    st.metric("Extracted input files", extracted_count)
    downloaded_count = len([p for p in st.session_state.get("downloaded_references", []) if p.get("download_success")])
    st.metric("Full-text references", downloaded_count)
    gemini_count = len([p for p in st.session_state.get("selected_papers", []) if p.get("gemini_note")])
    st.metric("Gemini-read references", gemini_count)


st.title("Research Paper Writing App")
st.caption("Choose an author writing style, search and read evidence, suggest missing support, draft in that style, and run editor audits.")

tabs = st.tabs(["Inputs", "Writing Style", "Evidence Search", "Gemini Reading", "Draft Preview", "Style Audit", "Export"])


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

        if inputs_for_style := current_input_values():
            if inputs_for_style["master_context"] or inputs_for_style["raw_methodology"] or extracted_files:
                st.divider()
                if st.button("Suggest best author style from current inputs", width="stretch"):
                    context_text = current_context_text(inputs_for_style, extracted_files)
                    with st.spinner("Comparing methodology and results with all author style reports..."):
                        st.session_state.style_recommendation = recommend_writing_style(
                            context_text,
                            st.session_state.openai_key,
                            st.session_state.model,
                        )
                    recommendation = st.session_state.style_recommendation.get("recommended_style", "")
                    if recommendation in STYLE_PRESETS:
                        st.session_state.active_style_name = recommendation
                        st.session_state.active_style_path = STYLE_PRESETS[recommendation]
                        try:
                            st.session_state.active_style_contract = load_writing_style_contract(
                                st.session_state.active_style_path,
                                st.session_state.openai_key,
                                st.session_state.model,
                                st.session_state.active_style_name,
                            )
                            st.session_state.active_style_report = st.session_state.active_style_contract.get("profile") or {}
                        except Exception as exc:
                            st.warning(f"Suggested style could not be loaded automatically: {exc}")
                    st.success(f"Suggested style: {recommendation or 'not available'}")
                recommendation = st.session_state.get("style_recommendation") or {}
                if recommendation.get("recommended_style"):
                    st.info(f"Current suggested style: {recommendation.get('recommended_style')}")


with tabs[1]:
    st.subheader("Author Writing Style")
    st.caption("Choose one attached style report or let the app suggest the best style from your methodology and results.")

    preset_names = list(STYLE_PRESETS.keys())
    current_name = st.session_state.get("active_style_name", preset_names[0])
    if current_name not in preset_names:
        current_name = preset_names[0]
    selected_style_name = st.selectbox(
        "Author style preset",
        options=preset_names,
        index=preset_names.index(current_name),
    )
    if selected_style_name != st.session_state.get("active_style_name"):
        st.session_state.active_style_name = selected_style_name
        st.session_state.active_style_path = STYLE_PRESETS[selected_style_name]

    st.session_state.active_style_path = st.text_input(
        "Style report path",
        value=st.session_state.active_style_path,
    )

    style_cols = st.columns(2)
    with style_cols[0]:
        load_clicked = st.button("Load selected style", type="primary", width="stretch")
    with style_cols[1]:
        suggest_clicked = st.button("Suggest best style from inputs", width="stretch")

    if load_clicked:
        try:
            with st.spinner("Reading DOCX and building the selected author style contract..."):
                contract = load_writing_style_contract(
                    st.session_state.active_style_path,
                    st.session_state.openai_key,
                    st.session_state.model,
                    st.session_state.active_style_name,
                )
            st.session_state.active_style_contract = contract
            st.session_state.active_style_report = contract.get("profile") or {}
            st.session_state.openai_style_audit = {}
            st.session_state.gemini_style_audit = {}
            st.session_state.docx_bytes = None
            st.success(f"Loaded {contract.get('characters', 0)} characters from {contract.get('author')}.")
        except Exception as exc:
            st.error(str(exc))

    if suggest_clicked:
        inputs = current_input_values()
        extracted_files = current_extracted_files()
        if not (inputs["master_context"] or inputs["raw_methodology"] or extracted_files):
            st.warning("Add methodology, master context, or result files first.")
        else:
            context_text = current_context_text(inputs, extracted_files)
            with st.spinner("Comparing your methodology and results with all author style reports..."):
                st.session_state.style_recommendation = recommend_writing_style(
                    context_text,
                    st.session_state.openai_key,
                    st.session_state.model,
                )
            recommendation = st.session_state.style_recommendation.get("recommended_style", "")
            if recommendation in STYLE_PRESETS:
                st.session_state.active_style_name = recommendation
                st.session_state.active_style_path = STYLE_PRESETS[recommendation]
                try:
                    st.session_state.active_style_contract = load_writing_style_contract(
                        st.session_state.active_style_path,
                        st.session_state.openai_key,
                        st.session_state.model,
                        st.session_state.active_style_name,
                    )
                    st.session_state.active_style_report = st.session_state.active_style_contract.get("profile") or {}
                    st.session_state.openai_style_audit = {}
                    st.session_state.gemini_style_audit = {}
                    st.session_state.docx_bytes = None
                except Exception as exc:
                    st.warning(f"Suggested style could not be loaded automatically: {exc}")
            st.success(f"Suggested style: {recommendation or 'not available'}")

    recommendation = st.session_state.get("style_recommendation") or {}
    if recommendation:
        st.divider()
        st.subheader("Style Suitability Suggestion")
        st.write(recommendation.get("reason", ""))
        ranked_styles = recommendation.get("ranked_styles") or []
        if ranked_styles:
            st.dataframe(pd.DataFrame(ranked_styles), width="stretch", hide_index=True)
        if recommendation.get("methodology_style_advice"):
            st.write("Methodology style advice")
            st.write(recommendation["methodology_style_advice"])
        if recommendation.get("results_style_advice"):
            st.write("Results style advice")
            st.write(recommendation["results_style_advice"])

    contract = current_style_contract()
    profile = current_style_profile()
    if not contract:
        st.info("Load a style report before Gemini reading and drafting for strict style control.")
    else:
        cols = st.columns(5)
        cols[0].metric("Style characters", contract.get("characters", 0))
        cols[1].metric("Author style", profile.get("author", st.session_state.active_style_name))
        cols[2].metric("Style roles", len(current_style_library()))
        cols[3].metric("Planning chars", contract.get("compressed_planning_characters", 0))
        cols[4].metric("Cache", contract.get("style_cache_status", "") or "new")

        st.write("Style summary")
        st.write(profile.get("overall_style", ""))

        section_rules = profile.get("section_rules") or {}
        if section_rules:
            st.dataframe(
                pd.DataFrame([{"section": key, "rule": value} for key, value in section_rules.items()]),
                width="stretch",
                hide_index=True,
            )

        with st.expander("Signature moves and phrases", expanded=False):
            st.write("Signature moves")
            for item in profile.get("signature_moves", []):
                st.write(item)
            st.write("Phrases to use")
            for item in profile.get("phrases_to_use", []):
                st.write(item)
            st.write("Phrases to avoid")
            for item in profile.get("phrases_to_avoid", []):
                st.write(item)

        checklist = profile.get("editor_checklist") or []
        if checklist:
            with st.expander("Editor checklist", expanded=True):
                for item in checklist:
                    st.write(item)

        with st.expander("Attached style report preview", expanded=False):
            st.text(truncate_text(contract.get("text", ""), 10000))


with tabs[2]:
    st.subheader("Reference Search, Sorting, Filtering, and PDF Reading")
    st.caption(
        "Searches Google Scholar, Semantic Scholar, OpenAlex, CORE, Google PDF results, ResearchGate public pages, thesis repositories, and Perplexity Sonar when keys are available."
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

    query_tab, finding_tab, selection_tab = st.tabs(
        ["Query Making", "Reference Finding", "Reference Selection & Reading"]
    )

    with query_tab:
        st.markdown("### Results-First Query Making")
        st.caption(
            "Write the Results first, build the Discussion framework, and prepare the style-aware search queries before reference searching."
        )
        if st.button("Analyze findings, write Results, and build Discussion Framework", type="primary", width="stretch"):
            if not (inputs["master_context"] or inputs["raw_methodology"] or extracted_files):
                st.warning("Add methodology, master context, or result files before analyzing.")
            else:
                with st.spinner(
                    "Analyzing findings, writing Results, building the Discussion framework, and asking Claude to plan the evidence search..."
                ):
                    analysis, result_text, queries = build_context_for_search(inputs, extracted_files)
                st.session_state.analysis = analysis
                st.session_state.queries = queries
                st.session_state.paper_search = {}
                st.session_state.selected_papers = []
                st.session_state.downloaded_references = []
                st.session_state.gemini_recommendations = {}
                st.session_state.followup_suggestions = {}
                st.success("Results draft, Discussion framework, and Claude evidence plan are ready below.")

        analysis_state = st.session_state.get("analysis") or {}
        if analysis_state:
            with st.expander("Research questions and result logic", expanded=True):
                if analysis_state.get("analysis_title"):
                    st.write("Working title")
                    st.write(analysis_state.get("analysis_title"))
                if analysis_state.get("objective"):
                    st.write("Objective")
                    st.write(analysis_state.get("objective"))
                for label, key in [
                    ("Treatments / variables", "treatments_variables"),
                    ("Major findings", "major_findings"),
                    ("Significant differences", "significant_differences"),
                    ("Patterns and trends", "result_patterns"),
                    ("Likely research questions", "research_questions"),
                    ("Discussion needs", "discussion_needs"),
                ]:
                    values = analysis_state.get(key) or []
                    if values:
                        st.write(label)
                        for value in values:
                            st.write(value)

            if analysis_state.get("results_draft") or analysis_state.get("results_draft_error"):
                with st.expander("Results draft from analyzed findings", expanded=True):
                    st.caption("Written first from uploaded tables, figures, methodology context, and extracted result text.")
                    if analysis_state.get("results_draft_error"):
                        st.warning(analysis_state.get("results_draft_error"))
                    if analysis_state.get("results_draft"):
                        st.markdown(analysis_state.get("results_draft"))

            discussion_framework = analysis_state.get("discussion_framework") or {}
            if discussion_framework:
                with st.expander("Discussion framework from Results", expanded=True):
                    provider = discussion_framework.get("framework_provider") or ""
                    model_used = discussion_framework.get("framework_model") or ""
                    if provider or model_used:
                        st.caption(f"Framework engine: {provider} {model_used}".strip())
                    if discussion_framework.get("claude_framework_error"):
                        st.warning(f"Claude framework warning: {discussion_framework.get('claude_framework_error')}")
                    if discussion_framework.get("style_priority"):
                        st.write("Style priority")
                        st.write(discussion_framework.get("style_priority"))
                    if discussion_framework.get("discussion_thesis"):
                        st.write("Discussion thesis")
                        st.write(discussion_framework.get("discussion_thesis"))
                    paragraph_framework = discussion_framework.get("paragraph_framework") or []
                    if paragraph_framework:
                        st.write("Paragraph framework")
                        for item in paragraph_framework:
                            st.write(item)
                    citation_strategy = discussion_framework.get("citation_strategy") or []
                    if citation_strategy:
                        st.write("Citation strategy")
                        for item in citation_strategy:
                            st.write(item)

            claude_plan = analysis_state.get("claude_discussion_search_plan") or {}
            if claude_plan:
                with st.expander("Claude discussion evidence plan", expanded=True):
                    provider = claude_plan.get("query_provider") or "Claude"
                    model_used = claude_plan.get("model_used") or st.session_state.claude_model
                    st.caption(
                        f"{provider} planned what paper types are needed for the Discussion, then the same search engines used those queries."
                    )
                    if model_used:
                        st.caption(f"Model: {model_used}")
                    if claude_plan.get("claude_query_error"):
                        st.warning(f"Claude query planning warning: {claude_plan.get('claude_query_error')}")
                    needed_types = claude_plan.get("needed_paper_types") or []
                    if needed_types:
                        st.write("Needed paper types")
                        for item in needed_types:
                            st.write(item)
                    rationales = claude_plan.get("query_rationale") or []
                    if rationales:
                        st.write("Why these queries were made")
                        for item in rationales:
                            st.write(item)
                    missing_questions = claude_plan.get("missing_evidence_questions") or []
                    if missing_questions:
                        st.write("Questions to answer through reading")
                        for item in missing_questions:
                            st.write(item)

        if st.session_state.get("queries"):
            with st.expander("Search queries from Results and Discussion plan", expanded=True):
                for query in st.session_state.queries:
                    st.code(query, language="text")
        else:
            st.info("Run the analysis above to write Results and create search queries.")

    with finding_tab:
        st.markdown("### Reference Finding")
        st.caption("Use the planned queries to search and rank research papers, review papers, and theses.")
        if st.session_state.get("queries"):
            with st.expander("Queries to search", expanded=True):
                for query in st.session_state.queries:
                    st.code(query, language="text")
            if st.button("Search papers using these planned queries", type="primary", width="stretch"):
                analysis = st.session_state.get("analysis") or {}
                context_text = current_context_text(inputs, extracted_files, analysis)
                with st.spinner("Searching, deduplicating, and scoring papers from the planned Discussion queries..."):
                    search_result = search_and_rank_papers(
                        queries=st.session_state.queries,
                        context_text=context_text,
                        semantic_key=st.session_state.semantic_key,
                        serpapi_key=st.session_state.serpapi_key,
                        core_key=st.session_state.core_key,
                        perplexity_key=st.session_state.perplexity_key,
                        perplexity_model=st.session_state.perplexity_model,
                        openai_key=st.session_state.openai_key,
                        model=st.session_state.model,
                        reference_count=st.session_state.reference_count,
                        per_query_limit=st.session_state.per_query_limit,
                        use_ai_scoring=st.session_state.use_ai_scoring,
                    )
                st.session_state.paper_search = search_result
                st.session_state.selected_papers = search_result.get("selected", [])
                st.session_state.downloaded_references = []
                st.session_state.gemini_recommendations = {}
                st.success(
                    f"Found {search_result.get('candidate_count', 0)} candidates, "
                    f"{search_result.get('deduped_count', 0)} after deduplication, "
                    f"{len(st.session_state.selected_papers)} selected."
                )
        else:
            st.info("First use the Query Making tab to analyze findings and create search queries.")

        search_result = st.session_state.get("paper_search") or {}
        deep_queries = search_result.get("deep_queries") or {}
        if deep_queries:
            with st.expander("Independent agri deep-search queries", expanded=False):
                for label, key in [
                    ("Journal/research article", "journal_queries"),
                    ("Thesis/dissertation", "thesis_queries"),
                    ("Review paper", "review_queries"),
                ]:
                    queries_for_layer = deep_queries.get(key) or []
                    if queries_for_layer:
                        st.markdown(f"**{label}**")
                        for query in queries_for_layer:
                            st.code(query, language="text")

        warnings = search_result.get("warnings", [])
        if warnings:
            with st.expander("Search warnings", expanded=False):
                for warning in warnings:
                    st.warning(warning)

        quota_status = search_result.get("quota_status") or {}
        if quota_status:
            selected_counts = quota_status.get("selected_counts", {})
            req_cols = st.columns(3)
            req_cols[0].metric("Selected research articles", f"{selected_counts.get('Research Article', 0)}/10")
            req_cols[1].metric("Selected theses", f"{selected_counts.get('Thesis', 0)}/3")
            req_cols[2].metric("Selected review papers", f"{selected_counts.get('Review Paper', 0)}/1")

        papers = search_result.get("papers", [])
        if papers:
            df = paper_rows(papers)
            category_counts = df["category"].value_counts().to_dict() if "category" in df else {}
            count_cols = st.columns(3)
            count_cols[0].metric("Research articles found", category_counts.get("Research Article", 0))
            count_cols[1].metric("Review papers found", category_counts.get("Review Paper", 0))
            count_cols[2].metric("Theses found", category_counts.get("Thesis", 0))
            st.info("Go to Reference Selection & Reading to choose sources separately by type.")

    with selection_tab:
        st.markdown("### Reference Selection and Reading")
        search_result = st.session_state.get("paper_search") or {}
        papers = search_result.get("papers", [])
        if not papers:
            st.info("Run Reference Finding first. Search results will appear here for separate selection.")
        else:
            df = paper_rows(papers)
            category_counts = df["category"].value_counts().to_dict() if "category" in df else {}
            count_cols = st.columns(3)
            count_cols[0].metric("Research articles", category_counts.get("Research Article", 0))
            count_cols[1].metric("Review papers", category_counts.get("Review Paper", 0))
            count_cols[2].metric("Theses", category_counts.get("Thesis", 0))
            st.caption(
                "Select sources separately by type. The app will merge the chosen research papers, review papers, and theses for downloading, Gemini reading, and drafting."
            )
            category_sections = [
                ("Research papers", "Research Article", "research_articles"),
                ("Review papers", "Review Paper", "review_papers"),
                ("Theses / dissertations", "Thesis", "theses"),
            ]
            known_categories = {category for _label, category, _key in category_sections}
            other_papers = [paper for paper in papers if paper.get("category", "Research Article") not in known_categories]
            if other_papers:
                category_sections.append(("Other sources", "__OTHER__", "other_sources"))

            edited_frames = []
            selection_tabs = st.tabs([label for label, _category, _key in category_sections])
            for tab, (label, category, key_suffix) in zip(selection_tabs, category_sections):
                with tab:
                    if category == "__OTHER__":
                        category_papers = other_papers
                    else:
                        category_papers = [paper for paper in papers if paper.get("category", "Research Article") == category]
                    selected_in_category = len([paper for paper in category_papers if paper.get("selected")])
                    st.caption(f"{len(category_papers)} found; {selected_in_category} currently selected.")
                    if category_papers:
                        edited_frames.append(paper_selection_editor(category_papers, f"evidence_select_{key_suffix}"))
                    else:
                        st.info(f"No {label.lower()} found yet.")

            if st.button("Save selected sources from all sections", width="stretch"):
                selected_ids = set()
                for edited in edited_frames:
                    selected_ids.update(edited.loc[edited["selected"] == True, "paper_id"].astype(str).tolist())
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
                selected_summary = pd.DataFrame(
                    [
                        {"category": category, "selected": len([p for p in selected if p.get("category") == category])}
                        for category in ["Research Article", "Review Paper", "Thesis"]
                    ]
                )
                st.success(f"Saved {len(selected)} selected source(s) from separate evidence sections.")
                st.dataframe(selected_summary, width="stretch", hide_index=True)

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
                    if st.session_state.gemini_key:
                        context_text = current_context_text(inputs, extracted_files)
                        with st.spinner("Google Gemini is reading papers, thesis literature reviews, and thesis references..."):
                            downloaded = gemini_read_selected_papers(
                                downloaded,
                                context_text,
                                st.session_state.gemini_key,
                                st.session_state.gemini_model,
                                current_style_profile(),
                            )
                            st.session_state.gemini_recommendations = summarize_gemini_reference_notes(
                                downloaded,
                                context_text,
                                st.session_state.gemini_key,
                                st.session_state.gemini_model,
                                current_style_profile(),
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
                    gemini_count = len([p for p in downloaded if p.get("gemini_note")])
                    if st.session_state.gemini_key:
                        st.success(f"Downloaded/read {success_count}/{len(downloaded)} PDF(s); Gemini read {gemini_count} selected source(s).")
                    else:
                        st.success(f"Downloaded/read {success_count}/{len(downloaded)} selected paper(s). Add a Gemini key to read and rank the sources.")

                downloaded_refs = st.session_state.get("downloaded_references", [])
                if downloaded_refs:
                    download_rows = [
                        {
                            "status": "downloaded" if item.get("download_success") else "not found",
                            "method": item.get("download_method", ""),
                            "text_status": item.get("text_status", ""),
                            "text_engine": item.get("text_extraction_method", ""),
                            "chars": item.get("full_text_chars", 0),
                            "size_mb": round((item.get("pdf_size_bytes", 0) or 0) / 1024 / 1024, 2),
                            "category": item.get("category", ""),
                            "gemini": item.get("gemini_read_status", ""),
                            "title": item.get("title", ""),
                            "pdf_path": item.get("pdf_path", ""),
                        }
                        for item in downloaded_refs
                    ]
                    st.dataframe(pd.DataFrame(download_rows), width="stretch", hide_index=True)

                with st.expander("APA 7 reference preview", expanded=False):
                    for paper in st.session_state.get("selected_papers", []):
                        if include_in_final_references(paper):
                            st.write(format_apa_reference(paper))
                    skipped_theses = [
                        paper for paper in st.session_state.get("selected_papers", [])
                        if (paper.get("category") == "Thesis" and not include_in_final_references(paper))
                    ]
                    if skipped_theses:
                        st.caption(f"{len(skipped_theses)} thesis source(s) are used for RoL/reference mining only and are not shown as final citations.")


with tabs[3]:
    st.subheader("Gemini Reading and Reference Matching")
    inputs = current_input_values()
    extracted_files = current_extracted_files()
    selected_papers = st.session_state.get("selected_papers", [])
    downloaded_refs = st.session_state.get("downloaded_references", [])

    read_cols = st.columns(4)
    read_cols[0].metric("Selected sources", len(selected_papers))
    read_cols[1].metric("Full-text PDFs", len([p for p in selected_papers if p.get("download_success")]))
    read_cols[2].metric("Theses", len([p for p in selected_papers if p.get("category") == "Thesis"]))
    read_cols[3].metric("Gemini-read", len([p for p in selected_papers if p.get("gemini_note")]))

    if not selected_papers:
        st.info("Search and select papers first.")
    else:
        if st.button("Read selected sources with Gemini", type="primary", width="stretch"):
            if not st.session_state.gemini_key:
                st.error("Enter a Google Gemini API key in the sidebar first.")
            else:
                context_text = current_context_text(inputs, extracted_files)
                source_papers = downloaded_refs or selected_papers
                with st.spinner("Gemini is reading abstracts, downloaded full papers, thesis literature reviews, and thesis references..."):
                    read_papers = gemini_read_selected_papers(
                        source_papers,
                        context_text,
                        st.session_state.gemini_key,
                        st.session_state.gemini_model,
                        current_style_profile(),
                    )
                    st.session_state.gemini_recommendations = summarize_gemini_reference_notes(
                        read_papers,
                        context_text,
                        st.session_state.gemini_key,
                        st.session_state.gemini_model,
                        current_style_profile(),
                    )
                by_id = {str(item.get("paper_id")): item for item in read_papers}
                st.session_state.selected_papers = [by_id.get(str(p.get("paper_id")), p) for p in selected_papers]
                st.session_state.downloaded_references = [by_id.get(str(p.get("paper_id")), p) for p in downloaded_refs] if downloaded_refs else read_papers
                if st.session_state.paper_search.get("papers"):
                    st.session_state.paper_search["papers"] = [
                        by_id.get(str(paper.get("paper_id")), paper) for paper in st.session_state.paper_search["papers"]
                    ]
                st.success(f"Gemini read {len(read_papers)} selected source(s).")

        recommendations = st.session_state.get("gemini_recommendations") or {}
        if recommendations:
            st.divider()
            st.subheader("Most Relatable References")
            best_rows = recommendations.get("best_references") or []
            if best_rows:
                st.dataframe(pd.DataFrame(best_rows), width="stretch", hide_index=True)

            review = recommendations.get("review_paper_to_use") or {}
            if review:
                with st.expander("Review paper to use", expanded=True):
                    st.write(review)

            review_insights = recommendations.get("review_discussion_insights") or []
            if review_insights:
                with st.expander("Review-paper insights for Discussion", expanded=True):
                    for insight in review_insights[:20]:
                        st.write(insight)

            review_reference_leads = recommendations.get("review_reference_leads") or []
            if review_reference_leads:
                with st.expander("Original references found inside review papers", expanded=True):
                    for lead in review_reference_leads[:30]:
                        st.write(lead)

            review_primary_leads = recommendations.get("review_primary_study_leads") or []
            if review_primary_leads:
                with st.expander("Primary-study leads extracted from review bibliographies", expanded=True):
                    for lead in review_primary_leads[:30]:
                        st.write(lead)
                review_queries = review_primary_study_search_queries(
                    review_primary_leads,
                    current_context_text(inputs, extracted_files),
                )
                if review_queries:
                    with st.expander("Primary-paper search queries from review bibliographies", expanded=False):
                        for query in review_queries:
                            st.code(query, language="text")
                    if st.button("Search original papers from review references now", type="primary", width="stretch"):
                        context_text = current_context_text(inputs, extracted_files)
                        with st.spinner("Searching original primary papers named or implied in review bibliographies..."):
                            extra_search = search_and_rank_papers(
                                queries=review_queries,
                                context_text=context_text,
                                semantic_key=st.session_state.semantic_key,
                                serpapi_key=st.session_state.serpapi_key,
                                core_key=st.session_state.core_key,
                                perplexity_key=st.session_state.perplexity_key,
                                perplexity_model=st.session_state.perplexity_model,
                                openai_key=st.session_state.openai_key,
                                model=st.session_state.model,
                                reference_count=st.session_state.reference_count,
                                per_query_limit=st.session_state.per_query_limit,
                                use_ai_scoring=st.session_state.use_ai_scoring,
                            )
                        st.session_state.paper_search = merge_search_results(
                            st.session_state.get("paper_search") or {},
                            extra_search,
                            st.session_state.reference_count,
                        )
                        st.session_state.selected_papers = st.session_state.paper_search.get("selected", [])
                        selected_ids = {str(item.get("paper_id")) for item in st.session_state.selected_papers}
                        st.session_state.downloaded_references = [
                            item for item in st.session_state.get("downloaded_references", [])
                            if str(item.get("paper_id")) in selected_ids
                        ]
                        st.success(
                            f"Merged review-bibliography primary-paper search. Selected {len(st.session_state.selected_papers)} sources."
                        )

            thesis_leads = recommendations.get("thesis_reference_leads") or []
            if thesis_leads:
                with st.expander("Useful references found inside theses", expanded=True):
                    for lead in thesis_leads[:30]:
                        st.write(lead)

            primary_leads = recommendations.get("thesis_primary_study_leads") or []
            if primary_leads:
                with st.expander("Primary-study leads extracted from thesis RoL", expanded=True):
                    for lead in primary_leads[:30]:
                        st.write(lead)
                rol_queries = thesis_primary_study_search_queries(
                    primary_leads,
                    current_context_text(inputs, extracted_files),
                )
                if rol_queries:
                    with st.expander("Primary-paper search queries from thesis RoL", expanded=False):
                        for query in rol_queries:
                            st.code(query, language="text")
                    if st.button("Search original papers from thesis RoL leads now", type="primary", width="stretch"):
                        context_text = current_context_text(inputs, extracted_files)
                        with st.spinner("Searching original primary papers named or implied in thesis RoL sections..."):
                            extra_search = search_and_rank_papers(
                                queries=rol_queries,
                                context_text=context_text,
                                semantic_key=st.session_state.semantic_key,
                                serpapi_key=st.session_state.serpapi_key,
                                core_key=st.session_state.core_key,
                                perplexity_key=st.session_state.perplexity_key,
                                perplexity_model=st.session_state.perplexity_model,
                                openai_key=st.session_state.openai_key,
                                model=st.session_state.model,
                                reference_count=st.session_state.reference_count,
                                per_query_limit=st.session_state.per_query_limit,
                                use_ai_scoring=st.session_state.use_ai_scoring,
                            )
                        st.session_state.paper_search = merge_search_results(
                            st.session_state.get("paper_search") or {},
                            extra_search,
                            st.session_state.reference_count,
                        )
                        st.session_state.selected_papers = st.session_state.paper_search.get("selected", [])
                        selected_ids = {str(item.get("paper_id")) for item in st.session_state.selected_papers}
                        st.session_state.downloaded_references = [
                            item for item in st.session_state.get("downloaded_references", [])
                            if str(item.get("paper_id")) in selected_ids
                        ]
                        st.success(
                            f"Merged thesis-RoL primary-paper search. Selected {len(st.session_state.selected_papers)} sources."
                        )

            gaps = recommendations.get("coverage_gaps") or []
            if gaps:
                with st.expander("Remaining literature gaps", expanded=False):
                    for gap in gaps:
                        st.write(gap)

        st.divider()
        st.subheader("Suggested Extra Evidence or Data")
        if st.button("Suggest missing papers, review support, thesis leads, or data checks", width="stretch"):
            if not (st.session_state.claude_key or st.session_state.gemini_key or st.session_state.openai_key):
                st.error("Enter a Claude, Gemini, or OpenAI key first.")
            else:
                context_text = current_context_text(inputs, extracted_files)
                with st.spinner("Checking what extra evidence or data would improve the selected-author-style paper..."):
                    st.session_state.followup_suggestions = suggest_style_aligned_followup_needs(
                        context_text=context_text,
                        selected_papers=st.session_state.get("selected_papers", []),
                        recommendations=st.session_state.get("gemini_recommendations", {}),
                        style_profile=current_style_profile(),
                        gemini_key=st.session_state.gemini_key,
                        gemini_model=st.session_state.gemini_model,
                        claude_key=st.session_state.claude_key,
                        claude_model=st.session_state.claude_model,
                        openai_key=st.session_state.openai_key,
                        openai_model=st.session_state.model,
                    )
                st.success("Suggestions prepared.")

        suggestions = st.session_state.get("followup_suggestions") or {}
        if suggestions:
            if suggestions.get("query_provider"):
                model_note = f" ({suggestions.get('model_used')})" if suggestions.get("model_used") else ""
                st.caption(f"Suggestion engine: {suggestions.get('query_provider')}{model_note}")
            if suggestions.get("claude_query_error"):
                st.warning(f"Claude follow-up planning warning: {suggestions.get('claude_query_error')}")
            search_queries = [
                clean_query_text(query)
                for query in suggestions.get("suggested_search_queries", [])
                if clean_query_text(query)
            ]
            if suggestions.get("style_plan"):
                with st.expander("Selected-style writing plan", expanded=True):
                    for item in suggestions.get("style_plan", []):
                        st.write(item)
            if suggestions.get("needed_data_checks"):
                with st.expander("Data or method details to provide if available", expanded=True):
                    for item in suggestions.get("needed_data_checks", []):
                        st.write(item)
            if search_queries:
                with st.expander("Suggested follow-up search queries", expanded=True):
                    for query in search_queries:
                        st.code(query, language="text")
                if st.button("Search suggested evidence now", type="primary", width="stretch"):
                    context_text = current_context_text(inputs, extracted_files)
                    with st.spinner("Searching suggested evidence with the same scholarly search engine..."):
                        extra_search = search_and_rank_papers(
                            queries=search_queries,
                            context_text=context_text,
                            semantic_key=st.session_state.semantic_key,
                            serpapi_key=st.session_state.serpapi_key,
                            core_key=st.session_state.core_key,
                            perplexity_key=st.session_state.perplexity_key,
                            perplexity_model=st.session_state.perplexity_model,
                            openai_key=st.session_state.openai_key,
                            model=st.session_state.model,
                            reference_count=st.session_state.reference_count,
                            per_query_limit=st.session_state.per_query_limit,
                            use_ai_scoring=st.session_state.use_ai_scoring,
                        )
                    st.session_state.paper_search = merge_search_results(
                        st.session_state.get("paper_search") or {},
                        extra_search,
                        st.session_state.reference_count,
                    )
                    st.session_state.selected_papers = st.session_state.paper_search.get("selected", [])
                    selected_ids = {str(item.get("paper_id")) for item in st.session_state.selected_papers}
                    st.session_state.downloaded_references = [
                        item for item in st.session_state.get("downloaded_references", []) if str(item.get("paper_id")) in selected_ids
                    ]
                    st.success(f"Merged suggested evidence search. Selected {len(st.session_state.selected_papers)} sources.")

        note_rows = []
        for paper in st.session_state.get("selected_papers", []):
            note = paper.get("gemini_note") or {}
            note_rows.append(
                {
                    "category": paper.get("category", ""),
                    "citation": citation_key(paper),
                    "gemini_status": paper.get("gemini_read_status", ""),
                    "relevance": note.get("overall_relevance", ""),
                    "evidence": note.get("evidence_source", ""),
                    "title": paper.get("title", ""),
                    "why_relevant": note.get("why_relevant", ""),
                    "style_use": note.get("selected_style_use") or note.get("shelton_style_use", ""),
                }
            )
        if note_rows:
            st.dataframe(pd.DataFrame(note_rows), width="stretch", hide_index=True)

        thesis_notes = [
            (paper, paper.get("gemini_note") or {})
            for paper in st.session_state.get("selected_papers", [])
            if paper.get("category") == "Thesis" and paper.get("gemini_note")
        ]
        if thesis_notes:
            with st.expander("Thesis literature-review and bibliography notes", expanded=False):
                for paper, note in thesis_notes:
                    st.markdown(f"**{citation_key(paper)} {paper.get('title', '')}**")
                    if note.get("review_of_literature_notes"):
                        st.write(note["review_of_literature_notes"])
                    for reference in note.get("most_useful_references", [])[:10]:
                        st.write(reference)


with tabs[4]:
    st.subheader("Draft Preview")
    inputs = current_input_values()
    extracted_files = current_extracted_files()
    selected_papers = st.session_state.get("selected_papers", [])

    cols = st.columns(5)
    cols[0].metric("Selected references", len(selected_papers))
    cols[1].metric("Full-text references", len([p for p in selected_papers if p.get("download_success")]))
    cols[2].metric("Gemini-read", len([p for p in selected_papers if p.get("gemini_note")]))
    cols[3].metric("Tables", len(collect_tables(extracted_files)))
    cols[4].metric("Graphs", len(collect_images(extracted_files)))
    st.caption(
        f"Discussion engine: {'Claude ' + st.session_state.claude_model if st.session_state.claude_key else 'OpenAI ' + st.session_state.model}"
    )

    with st.expander("Results-first discussion framework", expanded=False):
        st.markdown(DISCUSSION_WORKFLOW_TEXT)

    with st.expander("SAU/ICAR Results writing guide and reference style", expanded=False):
        st.markdown(truncate_text(load_sau_icar_results_prompt(), 12000))

    if st.button("Generate selected-style full paper draft", type="primary", width="stretch"):
        if not st.session_state.openai_key:
            st.error("Enter an OpenAI API key in the sidebar before drafting.")
        elif not (inputs["master_context"] or inputs["raw_methodology"] or extracted_files):
            st.warning("Add methodology, master context, or result files before drafting.")
        else:
            if not current_style_library():
                st.warning("Load a writing style report first for strict style control. Drafting will continue with only general prompts.")
            if selected_papers and not any(p.get("download_success") for p in selected_papers):
                st.warning("Selected references have not been downloaded/read yet; discussion will rely mostly on abstracts.")
            if selected_papers and not any(p.get("gemini_note") for p in selected_papers):
                st.warning("Run Gemini Reading first for stronger literature matching and thesis-reference extraction.")
            with st.spinner("Writing Results first, planning the style-led Discussion framework, then drafting the full paper..."):
                try:
                    draft = generate_full_draft(
                        api_key=st.session_state.openai_key,
                        model=st.session_state.model,
                        claude_key=st.session_state.claude_key,
                        claude_model=st.session_state.claude_model,
                        paper_title=inputs["paper_title"],
                        authors=inputs["authors"],
                        affiliation=inputs["affiliation"],
                        research_area=inputs["research_area"],
                        master_context=inputs["master_context"],
                        raw_methodology=inputs["raw_methodology"],
                        extracted_files=extracted_files,
                        styles=current_style_library(),
                        selected_papers=selected_papers,
                        writing_length_mode=st.session_state.writing_length_mode,
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
        if draft.get("discussion_model"):
            st.caption(f"Discussion written with {draft.get('discussion_provider', 'LLM')}: {draft['discussion_model']}")
        if draft.get("writing_length_mode"):
            st.caption(f"Writing length: {draft.get('writing_length_mode')}")
        section_map = [
            ("Abstract", "abstract"),
            ("Introduction", "introduction"),
            ("Materials and Methods", "methodology"),
            ("Results", "results"),
            ("Discussion Framework", "discussion_framework"),
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
                elif key == "discussion_framework":
                    st.json(value)
                else:
                    st.write(value)


with tabs[5]:
    st.subheader("Style Audit and Editor Check")
    draft = st.session_state.get("draft") or {}
    profile = current_style_profile()

    if not draft:
        st.info("Generate a draft before running the style audit.")
    elif not profile:
        st.warning("Load a writing style report first so the editors can audit against the exact style contract.")
    else:
        audit_cols = st.columns(3)
        audit_cols[0].metric("OpenAI audit", st.session_state.get("openai_style_audit", {}).get("score", "not run"))
        audit_cols[1].metric("Gemini final check", st.session_state.get("gemini_style_audit", {}).get("score", "not run"))
        audit_cols[2].metric("Style revision", "applied" if draft.get("style_revision_applied") else "not applied")

        left, right = st.columns(2)
        with left:
            if st.button("Run OpenAI style editor audit", type="primary", width="stretch"):
                if not st.session_state.openai_key:
                    st.error("Enter an OpenAI API key first.")
                else:
                    with st.spinner("OpenAI editor is auditing the draft against the selected style report..."):
                        st.session_state.openai_style_audit = audit_draft_with_openai_style_editor(
                            st.session_state.openai_key,
                            st.session_state.model,
                            draft,
                            profile,
                        )
                    st.success("OpenAI style audit complete.")
        with right:
            if st.button("Run Gemini final editor check", type="primary", width="stretch"):
                if not st.session_state.gemini_key:
                    st.error("Enter a Google Gemini API key first.")
                else:
                    with st.spinner("Gemini editor is running the final selected-style check..."):
                        st.session_state.gemini_style_audit = audit_draft_with_gemini_style_editor(
                            st.session_state.gemini_key,
                            st.session_state.gemini_model,
                            draft,
                            profile,
                        )
                    st.success("Gemini final editor check complete.")

        if st.button("Revise draft using style audits", width="stretch"):
            if not st.session_state.openai_key:
                st.error("Enter an OpenAI API key first.")
            elif not (st.session_state.get("openai_style_audit") or st.session_state.get("gemini_style_audit")):
                st.warning("Run at least one style audit before revision.")
            else:
                with st.spinner("Revising the draft while preserving facts, values, and citations..."):
                    st.session_state.draft = revise_draft_with_style_audits(
                        st.session_state.openai_key,
                        st.session_state.model,
                        draft,
                        profile,
                        st.session_state.get("openai_style_audit"),
                        st.session_state.get("gemini_style_audit"),
                    )
                    st.session_state.docx_bytes = None
                st.success("Draft revised using the style audits.")

        openai_audit = st.session_state.get("openai_style_audit") or {}
        if openai_audit:
            with st.expander("OpenAI editor audit", expanded=True):
                st.write(openai_audit)

        gemini_audit = st.session_state.get("gemini_style_audit") or {}
        if gemini_audit:
            with st.expander("Gemini final editor check", expanded=True):
                st.write(gemini_audit)


with tabs[6]:
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
