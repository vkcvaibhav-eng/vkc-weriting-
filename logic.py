from __future__ import annotations

import base64
import io
import json
import math
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
import urllib3
import httpx
from openai import OpenAI
from pypdf import PdfReader

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
except Exception:  # pragma: no cover - import error is shown in the Streamlit UI
    Document = None
    WD_ALIGN_PARAGRAPH = None
    Inches = None
    Pt = None


APP_ROOT = Path(__file__).resolve().parent
STYLE_LIBRARY_DIR = APP_ROOT / "style_library"
OUTPUT_DIR = APP_ROOT / "outputs"
DOWNLOADED_REFERENCES_DIR = APP_ROOT / "downloaded_references"
SSL_VERIFY = os.getenv("RP_APP_VERIFY_SSL", "false").strip().lower() in {"1", "true", "yes", "on"}

if not SSL_VERIFY:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

DEFAULT_STYLE_SOURCES = {
    "abstract": {
        "path": r"C:\Users\Admin\Downloads\Abstract .pdf",
        "role": "Field-study abstract style with objective, location, methods, results, and conclusion.",
    },
    "introduction": {
        "path": r"G:\My Drive\#NAU\Publication\LM\Research paper\2.Introduction\Introduction.pdf",
        "role": "Crop importance, pest/problem severity, knowledge gap, and objective style.",
    },
    "introduction_dictionary": {
        "path": r"G:\My Drive\#NAU\Publication\LM\Research paper\2.Introduction\Dictinery.pdf",
        "role": "Introduction connectors, technical agricultural vocabulary, and objective phrasing.",
    },
    "discussion": {
        "path": r"G:\My Drive\#NAU\Publication\LM\Research paper\3.Discussion\Discussion.pdf",
        "role": "Result-to-literature comparison, correlation interpretation, and cited support style.",
    },
    "methodology": {
        "path": r"G:\My Drive\#NAU\Publication\LM\Research paper\4. Matrial and methods\Matrial and methods.pdf",
        "role": "Passive factual methodology style: design, location, treatments, sampling, statistics.",
    },
    "word_bank": {
        "path": r"G:\My Drive\#NAU\Publication\LM\word used for described.pdf",
        "role": "Preferred statistical, treatment ranking, and comparative writing phrases.",
    },
    "input_template": {
        "path": r"G:\My Drive\#NAU\Publication\Online RPapp\Pre_App_Research_Paper_Input_Template.pdf",
        "role": "Author-ready input template and field structure.",
    },
}

STYLE_ROLE_LABELS = {
    "abstract": "Abstract",
    "introduction": "Introduction",
    "introduction_dictionary": "Introduction Dictionary",
    "discussion": "Discussion",
    "methodology": "Methodology",
    "word_bank": "Word Bank",
    "input_template": "Input Template",
}

AGRO_SANDESH_STYLE_GUIDE = """
Write as an experienced agricultural scientist and extension communicator inspired by
Dr. M. S. Swaminathan's farmer-first public communication style, without claiming
authorship by him. The article is for Agro Sandesh and must be in Gujarati.

Audience: farmers, extension workers, rural youth, agriculture students, and
progressive growers.

Structure:
1. Opening field problem or farmer observation.
2. Why the problem is increasing in the present season/month.
3. Scientific explanation in simple farmer-friendly language.
4. Practical field observations, preferably from Gujarat and South Gujarat when relevant.
5. Step-by-step integrated management recommendations.
6. Explain what to do, why to do it, and how it benefits the farmer for every recommendation.
7. Future opportunities, innovations, monitoring, or emerging technologies.
8. Positive practical takeaway.

Tone: farmer-centric, practical, trustworthy, inspirational, evidence-based,
solution-oriented, and scientifically accurate.

Avoid:
- Research paper, thesis, review article, or technical-report style.
- Headings like Introduction, Materials and Methods, Results, Discussion, Conclusion.
- Excessive jargon. Explain technical terms immediately in simple language.
- Emotional, political, or unsupported claims.

Connect science with crop health, yield, quality, cost reduction, profitability,
sustainability, timely monitoring, preventive action, integrated management, natural
enemies, and long-term benefit.
"""

GUJARAT_CROP_ACTIVITY_BY_MONTH = {
    "January": "rabi crop care, vegetable protection, mango flowering care, and monitoring of stored produce or nursery pests",
    "February": "late rabi monitoring, mango flowering and fruit-set care, vegetable pest management, and summer crop planning",
    "March": "summer crop sowing, mango fruit development, vegetable mite and sucking-pest monitoring, and irrigation-linked pest risk",
    "April": "summer crop stress management, mango fruit care, vegetable and nursery pest surveillance, and hot-weather mite risk",
    "May": "pre-monsoon field sanitation, nursery planning, mango harvest care, and preparation for kharif pest prevention",
    "June": "monsoon onset, kharif sowing, paddy nursery care, cotton and pulse establishment, and early sucking-pest monitoring",
    "July": "active kharif growth, rice/cotton/pulse/vegetable pest surveillance, weed and humidity-linked pest risk",
    "August": "peak monsoon crop growth, humid-weather pest and disease pressure, rice and cotton pest monitoring, and natural enemy conservation",
    "September": "late kharif crop protection, cotton/pulse/vegetable pest pressure, rice reproductive-stage monitoring, and harvest-quality protection",
    "October": "kharif harvest, rabi sowing preparation, vegetable nursery care, cotton picking-stage pest management, and residue sanitation",
    "November": "rabi establishment, vegetable and pulse pest monitoring, mango orchard sanitation, and winter pest prevention",
    "December": "rabi crop growth, vegetable pest management, mango orchard care, and cool-season pest monitoring",
}

STOPWORDS = {
    "about",
    "above",
    "after",
    "again",
    "against",
    "also",
    "among",
    "and",
    "are",
    "because",
    "been",
    "both",
    "but",
    "can",
    "crop",
    "data",
    "did",
    "during",
    "each",
    "effect",
    "from",
    "had",
    "has",
    "have",
    "into",
    "its",
    "may",
    "more",
    "not",
    "number",
    "observed",
    "over",
    "paper",
    "per",
    "plant",
    "population",
    "recorded",
    "research",
    "result",
    "results",
    "showed",
    "study",
    "than",
    "that",
    "the",
    "their",
    "these",
    "this",
    "through",
    "under",
    "used",
    "using",
    "was",
    "were",
    "with",
}


@dataclass
class ExtractedFile:
    name: str
    kind: str
    text: str
    tables: list[dict[str, Any]]
    image_bytes: bytes | None = None
    mime_type: str | None = None
    summary: str = ""


def ensure_app_dirs() -> None:
    STYLE_LIBRARY_DIR.mkdir(exist_ok=True)
    (STYLE_LIBRARY_DIR / "uploads").mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    DOWNLOADED_REFERENCES_DIR.mkdir(exist_ok=True)


def safe_slug(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value.strip("_") or "file"


def truncate_text(text: str, limit: int = 6000) -> str:
    clean = re.sub(r"\s+", " ", text or "").strip()
    if len(clean) <= limit:
        return clean
    return clean[:limit].rsplit(" ", 1)[0] + "..."


def get_llm_client(api_key: str, base_url: str | None = None) -> OpenAI:
    kwargs: dict[str, Any] = {"api_key": api_key}
    if base_url:
        kwargs["base_url"] = base_url
    if not SSL_VERIFY:
        kwargs["http_client"] = httpx.Client(verify=False, timeout=120)
    return OpenAI(**kwargs)


def chat_text(
    api_key: str,
    model: str,
    system: str,
    user: str,
    temperature: float = 0.25,
    response_format: dict[str, Any] | None = None,
) -> str:
    client = get_llm_client(api_key)
    kwargs: dict[str, Any] = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": temperature,
    }
    if response_format:
        kwargs["response_format"] = response_format
    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content or ""


def parse_json_object(text: str, fallback: dict[str, Any]) -> dict[str, Any]:
    if not text:
        return fallback
    candidate = text.strip()
    candidate = re.sub(r"^```(?:json)?", "", candidate, flags=re.IGNORECASE).strip()
    candidate = re.sub(r"```$", "", candidate).strip()
    try:
        return json.loads(candidate)
    except Exception:
        pass
    match = re.search(r"\{.*\}", candidate, flags=re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            return fallback
    return fallback


def parse_json_list(text: str, fallback: list[Any]) -> list[Any]:
    if not text:
        return fallback
    candidate = text.strip()
    candidate = re.sub(r"^```(?:json)?", "", candidate, flags=re.IGNORECASE).strip()
    candidate = re.sub(r"```$", "", candidate).strip()
    try:
        parsed = json.loads(candidate)
        return parsed if isinstance(parsed, list) else fallback
    except Exception:
        pass
    match = re.search(r"\[.*\]", candidate, flags=re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group(0))
            return parsed if isinstance(parsed, list) else fallback
        except Exception:
            return fallback
    return fallback


def extract_text_from_pdf(source: str | Path | io.BytesIO) -> str:
    reader = PdfReader(source)
    parts: list[str] = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            continue
    return "\n".join(parts).strip()


def extract_docx(source: str | Path | io.BytesIO) -> tuple[str, list[dict[str, Any]]]:
    if Document is None:
        raise RuntimeError("python-docx is required to read DOCX files.")
    document = Document(source)
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
    tables: list[dict[str, Any]] = []
    for index, table in enumerate(document.tables, start=1):
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        if rows:
            tables.append({"name": f"DOCX Table {index}", "rows": rows})
    table_text = "\n".join(table_to_plain_text(t) for t in tables)
    return "\n\n".join(paragraphs + ([table_text] if table_text else [])), tables


def table_to_plain_text(table: dict[str, Any], max_rows: int = 25) -> str:
    rows = table.get("rows") or []
    if not rows:
        return ""
    rendered = [table.get("name", "Table")]
    for row in rows[:max_rows]:
        rendered.append(" | ".join(str(cell) for cell in row))
    if len(rows) > max_rows:
        rendered.append(f"... {len(rows) - max_rows} additional rows")
    return "\n".join(rendered)


def dataframe_to_table(name: str, df: pd.DataFrame, max_rows: int = 100) -> dict[str, Any]:
    clean = df.fillna("")
    rows = [list(map(str, clean.columns.tolist()))]
    rows.extend(clean.astype(str).head(max_rows).values.tolist())
    return {"name": name, "rows": rows, "shape": clean.shape}


def extract_spreadsheet(uploaded_bytes: bytes, filename: str) -> tuple[str, list[dict[str, Any]]]:
    suffix = Path(filename).suffix.lower()
    tables: list[dict[str, Any]] = []
    if suffix == ".csv":
        df = pd.read_csv(io.BytesIO(uploaded_bytes))
        tables.append(dataframe_to_table(filename, df))
    elif suffix in {".xlsx", ".xls"}:
        sheets = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=None)
        for sheet_name, df in sheets.items():
            tables.append(dataframe_to_table(f"{filename} - {sheet_name}", df))
    else:
        raise ValueError(f"Unsupported spreadsheet type: {suffix}")
    text = "\n\n".join(table_to_plain_text(table) for table in tables)
    return text, tables


def extract_uploaded_file(uploaded_file: Any, openai_key: str = "", model: str = "gpt-4o-mini") -> ExtractedFile:
    filename = uploaded_file.name
    suffix = Path(filename).suffix.lower()
    data = uploaded_file.getvalue()
    mime_type = getattr(uploaded_file, "type", None) or ""

    if suffix == ".pdf":
        text = extract_text_from_pdf(io.BytesIO(data))
        return ExtractedFile(filename, "pdf", text, [], mime_type=mime_type)
    if suffix == ".docx":
        text, tables = extract_docx(io.BytesIO(data))
        return ExtractedFile(filename, "docx", text, tables, mime_type=mime_type)
    if suffix in {".csv", ".xlsx", ".xls"}:
        text, tables = extract_spreadsheet(data, filename)
        return ExtractedFile(filename, "table", text, tables, mime_type=mime_type)
    if suffix in {".txt", ".md"}:
        text = data.decode("utf-8", errors="ignore")
        return ExtractedFile(filename, "text", text, [], mime_type=mime_type)
    if suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        summary = summarize_image_with_llm(data, mime_type or "image/png", openai_key, model) if openai_key else ""
        return ExtractedFile(filename, "image", summary, [], image_bytes=data, mime_type=mime_type, summary=summary)
    return ExtractedFile(filename, "unknown", "", [], mime_type=mime_type)


def summarize_image_with_llm(image_bytes: bytes, mime_type: str, api_key: str, model: str) -> str:
    client = get_llm_client(api_key)
    data_url = f"data:{mime_type};base64,{base64.b64encode(image_bytes).decode('ascii')}"
    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Interpret this research result graph or image. Extract axes, treatments, "
                            "important trends, significant differences if shown, and a concise result statement. "
                            "Do not invent values that are not visible."
                        ),
                    },
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            }
        ],
        temperature=0.2,
    )
    return response.choices[0].message.content or ""


def extracted_files_to_dict(files: list[ExtractedFile]) -> list[dict[str, Any]]:
    return [
        {
            "name": item.name,
            "kind": item.kind,
            "text": item.text,
            "tables": item.tables,
            "image_bytes": item.image_bytes,
            "mime_type": item.mime_type,
            "summary": item.summary,
        }
        for item in files
    ]


def extracted_files_from_dict(items: list[dict[str, Any]]) -> list[ExtractedFile]:
    return [
        ExtractedFile(
            name=item.get("name", ""),
            kind=item.get("kind", ""),
            text=item.get("text", ""),
            tables=item.get("tables", []),
            image_bytes=item.get("image_bytes"),
            mime_type=item.get("mime_type"),
            summary=item.get("summary", ""),
        )
        for item in items
    ]


def cache_default_style_sources() -> list[dict[str, Any]]:
    ensure_app_dirs()
    statuses = []
    for role, meta in DEFAULT_STYLE_SOURCES.items():
        source_path = Path(meta["path"])
        cache_path = STYLE_LIBRARY_DIR / f"default_{role}.txt"
        status = {
            "role": role,
            "label": STYLE_ROLE_LABELS.get(role, role),
            "source": str(source_path),
            "cache": str(cache_path),
            "exists": source_path.exists(),
            "characters": 0,
            "status": "missing",
        }
        if source_path.exists():
            try:
                text = extract_text_from_pdf(source_path)
                cache_path.write_text(text, encoding="utf-8")
                status["characters"] = len(text)
                status["status"] = "loaded"
            except Exception as exc:
                status["status"] = f"error: {exc}"
        statuses.append(status)
    return statuses


def load_style_library(include_defaults: bool = True) -> tuple[dict[str, str], list[dict[str, Any]]]:
    ensure_app_dirs()
    statuses: list[dict[str, Any]] = []
    if include_defaults:
        statuses = cache_default_style_sources()

    styles: dict[str, list[str]] = {role: [] for role in STYLE_ROLE_LABELS}
    for text_path in STYLE_LIBRARY_DIR.glob("*.txt"):
        role = text_path.stem.replace("default_", "", 1)
        if role in styles:
            styles[role].append(text_path.read_text(encoding="utf-8", errors="ignore"))

    for text_path in (STYLE_LIBRARY_DIR / "uploads").glob("*.txt"):
        role = text_path.name.split("__", 1)[0]
        if role in styles:
            styles[role].append(text_path.read_text(encoding="utf-8", errors="ignore"))

    merged = {role: "\n\n".join(parts).strip() for role, parts in styles.items()}
    return merged, statuses


def save_uploaded_style(uploaded_file: Any, role: str) -> dict[str, Any]:
    ensure_app_dirs()
    if role not in STYLE_ROLE_LABELS:
        raise ValueError(f"Unknown style role: {role}")
    filename = safe_slug(uploaded_file.name)
    data = uploaded_file.getvalue()
    target = STYLE_LIBRARY_DIR / "uploads" / f"{role}__{filename}"
    target.write_bytes(data)
    suffix = target.suffix.lower()
    if suffix == ".pdf":
        text = extract_text_from_pdf(io.BytesIO(data))
    elif suffix == ".docx":
        text, _ = extract_docx(io.BytesIO(data))
    elif suffix in {".txt", ".md"}:
        text = data.decode("utf-8", errors="ignore")
    else:
        text = ""
    text_path = STYLE_LIBRARY_DIR / "uploads" / f"{role}__{filename}.txt"
    text_path.write_text(text, encoding="utf-8")
    return {"role": role, "file": str(target), "text_file": str(text_path), "characters": len(text)}


def style_excerpt(styles: dict[str, str], role: str, limit: int = 5000) -> str:
    return truncate_text(styles.get(role, ""), limit)


def combined_uploaded_text(files: list[ExtractedFile], limit: int = 18000) -> str:
    parts: list[str] = []
    for item in files:
        text = item.text or item.summary
        if text:
            parts.append(f"File: {item.name}\n{text}")
    return truncate_text("\n\n".join(parts), limit)


def collect_tables(files: list[ExtractedFile]) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for item in files:
        for table in item.tables:
            named = dict(table)
            named["source_file"] = item.name
            tables.append(named)
    return tables


def collect_images(files: list[ExtractedFile]) -> list[dict[str, Any]]:
    images = []
    for item in files:
        if item.image_bytes:
            images.append(
                {
                    "name": item.name,
                    "bytes": item.image_bytes,
                    "mime_type": item.mime_type,
                    "summary": item.summary or item.text,
                }
            )
    return images


def fallback_keywords(text: str, limit: int = 12) -> list[str]:
    words = re.findall(r"[A-Za-z][A-Za-z-]{3,}", text.lower())
    counts: dict[str, int] = {}
    for word in words:
        if word in STOPWORDS:
            continue
        counts[word] = counts.get(word, 0) + 1
    return [word for word, _ in sorted(counts.items(), key=lambda item: item[1], reverse=True)[:limit]]


def analyze_research_context(
    api_key: str,
    model: str,
    paper_title: str,
    research_area: str,
    master_context: str,
    raw_methodology: str,
    result_text: str,
) -> dict[str, Any]:
    fallback_title = paper_title.strip() or "Research Paper Draft"
    fallback = {
        "generated_title": fallback_title,
        "objective": "",
        "keywords": fallback_keywords(" ".join([paper_title, research_area, master_context, raw_methodology, result_text])),
        "major_findings": [],
        "treatment_rankings": [],
        "search_queries": [],
        "result_summary": truncate_text(result_text, 1200),
    }
    if not api_key:
        return fallback

    prompt = f"""
Return a JSON object for a research paper writing app.

Known title: {paper_title or "not supplied"}
Research area: {research_area or "not supplied"}
Master context:
{truncate_text(master_context, 5000)}

Raw methodology:
{truncate_text(raw_methodology, 7000)}

Result evidence:
{truncate_text(result_text, 9000)}

JSON keys:
generated_title, objective, keywords, major_findings, treatment_rankings, search_queries, result_summary.
Use only supplied facts. If title is missing, create a concise scientific title.
Create 5 to 8 search_queries that will find papers useful for discussion and references.
"""
    try:
        text = chat_text(
            api_key,
            model,
            "You extract structured context for agricultural and biological research papers.",
            prompt,
            temperature=0.1,
            response_format={"type": "json_object"},
        )
        parsed = parse_json_object(text, fallback)
        for key, value in fallback.items():
            parsed.setdefault(key, value)
        return parsed
    except Exception:
        return fallback


def generate_search_queries(
    api_key: str,
    model: str,
    analysis: dict[str, Any],
    master_context: str,
    result_text: str,
    max_queries: int = 6,
) -> list[str]:
    existing = [str(q).strip() for q in analysis.get("search_queries", []) if str(q).strip()]
    if existing:
        return existing[:max_queries]

    keywords = fallback_keywords(" ".join([master_context, result_text, " ".join(analysis.get("keywords", []))]))
    fallback = [
        " ".join(keywords[:5]) + " field study",
        " ".join(keywords[:4]) + " population dynamics",
        " ".join(keywords[:4]) + " treatment efficacy",
        " ".join(keywords[:4]) + " integrated pest management",
    ]
    fallback = [query.strip() for query in fallback if query.strip()]
    if not api_key:
        return fallback[:max_queries]

    prompt = f"""
Create search queries for finding research papers to support the discussion section.

Objective: {analysis.get("objective", "")}
Findings: {analysis.get("major_findings", [])}
Keywords: {analysis.get("keywords", [])}
Context: {truncate_text(master_context, 3000)}
Results: {truncate_text(result_text, 5000)}

Return JSON array of {max_queries} short queries. Include crop/host, pest/pathogen, treatment/weather/statistical terms when known.
"""
    try:
        text = chat_text(
            api_key,
            model,
            "You create precise scholarly search queries.",
            prompt,
            temperature=0.2,
        )
        queries = [str(q).strip() for q in parse_json_list(text, fallback) if str(q).strip()]
        return queries[:max_queries] or fallback[:max_queries]
    except Exception:
        return fallback[:max_queries]


def normalize_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (title or "").lower()).strip()


def infer_paper_category(paper: dict[str, Any] | None = None, title: str = "", source: str = "") -> str:
    paper = paper or {}
    title_text = (title or paper.get("title") or "").lower()
    source_text = (source or paper.get("source") or "").lower()
    pub_types = " ".join(map(str, paper.get("publication_types") or paper.get("publicationTypes") or [])).lower()
    if "thesis" in source_text or "krishikosh" in source_text or "thesis" in title_text or "dissertation" in title_text:
        return "Thesis"
    if "review" in source_text or "review" in title_text or "review" in pub_types:
        return "Review Paper"
    return "Research Article"


def author_names_from_semantic(authors: list[dict[str, Any]]) -> list[str]:
    names = []
    for author in authors or []:
        if isinstance(author, dict) and author.get("name"):
            names.append(str(author["name"]))
    return names


def search_semantic_scholar(query: str, api_key: str = "", limit: int = 10) -> list[dict[str, Any]]:
    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    params = {
        "query": query,
        "limit": limit,
        "fields": "title,authors,year,abstract,url,venue,citationCount,externalIds,publicationTypes,openAccessPdf,isOpenAccess",
    }
    headers = {"x-api-key": api_key} if api_key else {}
    response = requests.get(url, params=params, headers=headers, timeout=20, verify=SSL_VERIFY)
    response.raise_for_status()
    data = response.json()
    papers = []
    for item in data.get("data", []):
        papers.append(
            {
                "paper_id": item.get("paperId") or normalize_title(item.get("title", "")),
                "title": item.get("title") or "",
                "authors": author_names_from_semantic(item.get("authors", [])),
                "year": item.get("year"),
                "abstract": item.get("abstract") or "",
                "venue": item.get("venue") or "",
                "url": item.get("url") or "",
                "citation_count": item.get("citationCount") or 0,
                "source": "Semantic Scholar",
                "query": query,
                "external_ids": item.get("externalIds") or {},
                "pdf_url": (item.get("openAccessPdf") or {}).get("url") or "",
                "is_open_access": bool(item.get("isOpenAccess")),
                "publication_types": item.get("publicationTypes") or [],
                "category": infer_paper_category(
                    item,
                    title=item.get("title") or "",
                    source="Semantic Scholar",
                ),
            }
        )
    return papers


def search_serpapi_google_scholar(query: str, api_key: str, limit: int = 10) -> list[dict[str, Any]]:
    if not api_key:
        return []
    response = requests.get(
        "https://serpapi.com/search.json",
        params={"engine": "google_scholar", "q": query, "api_key": api_key, "num": limit},
        timeout=25,
        verify=SSL_VERIFY,
    )
    response.raise_for_status()
    data = response.json()
    papers = []
    for index, item in enumerate(data.get("organic_results", [])[:limit], start=1):
        publication_info = item.get("publication_info") or {}
        summary = publication_info.get("summary") or ""
        year_match = re.search(r"(19|20)\d{2}", summary)
        cited_by = ((item.get("inline_links") or {}).get("cited_by") or {}).get("total")
        pdf_urls = []
        for resource in item.get("resources", []) or []:
            if resource.get("file_format") == "PDF" and resource.get("link"):
                pdf_urls.append(resource["link"])
        papers.append(
            {
                "paper_id": f"serpapi-{normalize_title(item.get('title', ''))[:80]}-{index}",
                "title": item.get("title") or "",
                "authors": parse_author_summary(summary),
                "year": int(year_match.group(0)) if year_match else None,
                "abstract": item.get("snippet") or "",
                "venue": summary,
                "url": item.get("link") or "",
                "citation_count": cited_by or 0,
                "source": "SerpAPI Google Scholar",
                "query": query,
                "external_ids": {},
                "pdf_url": pdf_urls[0] if pdf_urls else "",
                "pdf_urls": pdf_urls,
                "cluster_id": publication_info.get("cites_id") or item.get("cluster_id") or "",
                "category": infer_paper_category(title=item.get("title") or "", source="SerpAPI Google Scholar"),
            }
        )
    return papers


def search_serpapi_review_layer(query: str, api_key: str, limit: int = 8) -> list[dict[str, Any]]:
    if not api_key:
        return []
    review_query = query if "review" in query.lower() else f'{query} review "state of the art"'
    papers = search_serpapi_google_scholar(review_query, api_key, limit)
    for paper in papers:
        paper["source"] = "Review Paper (Scholar)"
        paper["category"] = "Review Paper"
    return papers


def search_krishikosh_layer(query: str, api_key: str, limit: int = 8) -> list[dict[str, Any]]:
    if not api_key:
        return []
    full_query = query if "site:krishikosh" in query.lower() else f"{query} site:krishikosh.egranth.ac.in"
    response = requests.get(
        "https://serpapi.com/search.json",
        params={"engine": "google", "q": full_query, "api_key": api_key, "num": limit},
        timeout=25,
        verify=SSL_VERIFY,
    )
    response.raise_for_status()
    data = response.json()
    papers = []
    for index, item in enumerate(data.get("organic_results", [])[:limit], start=1):
        papers.append(
            {
                "paper_id": f"krishikosh-{normalize_title(item.get('title', ''))[:80]}-{index}",
                "title": item.get("title") or "",
                "authors": [],
                "year": None,
                "abstract": item.get("snippet") or "",
                "venue": "KrishiKosh",
                "url": item.get("link") or "",
                "citation_count": 0,
                "source": "KrishiKosh Thesis",
                "query": full_query,
                "external_ids": {},
                "pdf_url": item.get("link") if str(item.get("link", "")).lower().endswith(".pdf") else "",
                "category": "Thesis",
            }
        )
    return papers


def search_openalex(query: str, limit: int = 10) -> list[dict[str, Any]]:
    response = requests.get(
        "https://api.openalex.org/works",
        params={"search": query, "per-page": limit},
        timeout=25,
        verify=SSL_VERIFY,
    )
    response.raise_for_status()
    data = response.json()
    papers = []
    for index, item in enumerate(data.get("results", [])[:limit], start=1):
        authors = []
        for authorship in item.get("authorships", []) or []:
            author = authorship.get("author") or {}
            if author.get("display_name"):
                authors.append(author["display_name"])
        year = item.get("publication_year")
        doi = item.get("doi") or ""
        oa = item.get("open_access") or {}
        primary = item.get("primary_location") or {}
        pdf_url = oa.get("oa_url") or ((primary.get("source") or {}).get("pdf_url") or "")
        title = item.get("display_name") or ""
        papers.append(
            {
                "paper_id": item.get("id") or f"openalex-{index}-{normalize_title(title)[:80]}",
                "title": title,
                "authors": authors,
                "year": year,
                "abstract": openalex_abstract_to_text(item.get("abstract_inverted_index")),
                "venue": ((primary.get("source") or {}).get("display_name") or ""),
                "url": doi or item.get("id") or "",
                "citation_count": item.get("cited_by_count") or 0,
                "source": "OpenAlex",
                "query": query,
                "external_ids": {"DOI": doi.replace("https://doi.org/", "")} if doi else {},
                "pdf_url": pdf_url or "",
                "is_open_access": bool(oa.get("is_oa")),
                "category": infer_paper_category(title=title, source="OpenAlex"),
            }
        )
    return papers


def openalex_abstract_to_text(inverted_index: dict[str, list[int]] | None) -> str:
    if not inverted_index:
        return ""
    positioned = []
    for word, positions in inverted_index.items():
        for position in positions:
            positioned.append((position, word))
    return " ".join(word for _, word in sorted(positioned))


def search_core(query: str, api_key: str, limit: int = 10) -> list[dict[str, Any]]:
    if not api_key:
        return []
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        response = requests.get(
            "https://api.core.ac.uk/v3/search/works",
            params={"q": query, "limit": limit},
            headers=headers,
            timeout=25,
            verify=SSL_VERIFY,
        )
        response.raise_for_status()
    except Exception:
        response = requests.post(
            "https://api.core.ac.uk/v3/search/works",
            json={"q": query, "limit": limit},
            headers=headers,
            timeout=25,
            verify=SSL_VERIFY,
        )
    response.raise_for_status()
    data = response.json()
    papers = []
    for index, item in enumerate(data.get("results", [])[:limit], start=1):
        authors = []
        for author in item.get("authors", []) or []:
            if isinstance(author, dict):
                name = author.get("name") or author.get("fullName")
                if name:
                    authors.append(name)
            elif isinstance(author, str):
                authors.append(author)
        year = item.get("yearPublished") or item.get("publishedDate")
        if isinstance(year, str):
            year_match = re.search(r"(19|20)\d{2}", year)
            year = int(year_match.group(0)) if year_match else None
        papers.append(
            {
                "paper_id": str(item.get("id") or f"core-{index}-{normalize_title(item.get('title', ''))[:80]}"),
                "title": item.get("title") or "",
                "authors": authors,
                "year": year,
                "abstract": item.get("abstract") or "",
                "venue": item.get("publisher") or "",
                "url": item.get("downloadUrl") or item.get("doi") or "",
                "citation_count": item.get("citationCount") or 0,
                "source": "CORE",
                "query": query,
                "external_ids": {"DOI": item.get("doi")} if item.get("doi") else {},
                "pdf_url": item.get("downloadUrl") or "",
                "category": infer_paper_category(title=item.get("title") or "", source="CORE"),
            }
        )
    return papers


def parse_author_summary(summary: str) -> list[str]:
    if not summary:
        return []
    first = summary.split("-")[0].strip()
    first = re.sub(r"\s+", " ", first)
    names = [name.strip() for name in re.split(r",| and ", first) if name.strip()]
    return names[:8]


def deduplicate_papers(papers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_title: dict[str, dict[str, Any]] = {}
    for paper in papers:
        title_key = normalize_title(paper.get("title", ""))
        if not title_key:
            continue
        existing = by_title.get(title_key)
        if not existing:
            by_title[title_key] = paper
            continue
        merged = merge_paper_metadata(existing, paper)
        if int(paper.get("citation_count") or 0) > int(existing.get("citation_count") or 0):
            merged = merge_paper_metadata(paper, existing)
        by_title[title_key] = merged
    return list(by_title.values())


def merge_paper_metadata(primary: dict[str, Any], secondary: dict[str, Any]) -> dict[str, Any]:
    merged = dict(primary)
    for key in ["abstract", "url", "venue", "pdf_url", "year"]:
        if not merged.get(key) and secondary.get(key):
            merged[key] = secondary[key]
    if not merged.get("authors") and secondary.get("authors"):
        merged["authors"] = secondary["authors"]
    merged["citation_count"] = max(int(primary.get("citation_count") or 0), int(secondary.get("citation_count") or 0))
    sources = set(str(merged.get("source", "")).split(" + "))
    sources.add(str(secondary.get("source", "")))
    merged["source"] = " + ".join(sorted(source for source in sources if source))
    pdf_urls = set(merged.get("pdf_urls") or [])
    if merged.get("pdf_url"):
        pdf_urls.add(merged["pdf_url"])
    if secondary.get("pdf_url"):
        pdf_urls.add(secondary["pdf_url"])
    for url in secondary.get("pdf_urls") or []:
        pdf_urls.add(url)
    merged["pdf_urls"] = list(pdf_urls)
    if merged.get("category") == "Research Article" and secondary.get("category") in {"Review Paper", "Thesis"}:
        merged["category"] = secondary["category"]
    if not merged.get("external_ids"):
        merged["external_ids"] = secondary.get("external_ids") or {}
    else:
        merged["external_ids"] = {**(secondary.get("external_ids") or {}), **merged["external_ids"]}
    return merged


def heuristic_score_paper(paper: dict[str, Any], context_text: str) -> tuple[float, str]:
    keywords = set(fallback_keywords(context_text, 40))
    haystack = f"{paper.get('title', '')} {paper.get('abstract', '')}".lower()
    overlap = sum(1 for keyword in keywords if keyword in haystack)
    overlap_score = min(55, overlap * 5)

    year = paper.get("year")
    try:
        year_value = int(year)
    except Exception:
        year_value = 0
    if year_value >= 2020:
        year_score = 20
    elif year_value >= 2015:
        year_score = 14
    elif year_value >= 2005:
        year_score = 8
    elif year_value:
        year_score = 4
    else:
        year_score = 0

    citations = int(paper.get("citation_count") or 0)
    citation_score = min(20, math.log1p(citations) * 4)
    source_score = 5 if any(src in str(paper.get("source", "")) for src in ["Semantic Scholar", "CORE", "OpenAlex"]) else 3
    score = round(min(100, overlap_score + year_score + citation_score + source_score), 1)
    reason = f"keyword overlap {overlap}; year {year_value or 'unknown'}; citations {citations}"
    return score, reason


def ai_score_papers(
    api_key: str,
    model: str,
    papers: list[dict[str, Any]],
    context_text: str,
) -> list[dict[str, Any]]:
    if not api_key or not papers:
        return papers
    compact_papers = []
    for paper in papers[:30]:
        compact_papers.append(
            {
                "paper_id": paper.get("paper_id"),
                "title": paper.get("title"),
                "year": paper.get("year"),
                "abstract": truncate_text(paper.get("abstract", ""), 700),
                "source": paper.get("source"),
                "category": paper.get("category") or infer_paper_category(paper),
                "citation_count": paper.get("citation_count"),
                "has_pdf_link": bool(paper.get("pdf_url") or paper.get("pdf_urls")),
            }
        )
    prompt = f"""
Sort and classify each paper for usefulness in the Results and Discussion section of this research article.
Use relevance to the supplied crop/pest/treatment/weather/findings first, then recency and citation strength.

Research context and findings:
{truncate_text(context_text, 7000)}

Papers:
{json.dumps(compact_papers, ensure_ascii=True)}

Return JSON array. Each item must have paper_id, ai_score from 0 to 100, category as one of
"Research Article", "Review Paper", or "Thesis", and reason.
"""
    try:
        text = chat_text(
            api_key,
            model,
            "You select the best papers for scientific discussion. You do not invent metadata.",
            prompt,
            temperature=0.1,
        )
        scored = parse_json_list(text, [])
        score_map = {
            str(item.get("paper_id")): item
            for item in scored
            if isinstance(item, dict) and item.get("paper_id") is not None
        }
        for paper in papers:
            item = score_map.get(str(paper.get("paper_id")))
            if item:
                paper["ai_score"] = float(item.get("ai_score") or paper.get("score") or 0)
                paper["score_reason"] = item.get("reason") or paper.get("score_reason", "")
                paper["category"] = item.get("category") or paper.get("category") or infer_paper_category(paper)
                paper["score"] = round((float(paper.get("score") or 0) * 0.35) + (paper["ai_score"] * 0.65), 1)
            else:
                paper["category"] = paper.get("category") or infer_paper_category(paper)
        return papers
    except Exception:
        return papers


def search_and_rank_papers(
    queries: list[str],
    context_text: str,
    semantic_key: str = "",
    serpapi_key: str = "",
    core_key: str = "",
    openai_key: str = "",
    model: str = "gpt-4o-mini",
    reference_count: int = 12,
    per_query_limit: int = 8,
    use_ai_scoring: bool = True,
) -> dict[str, Any]:
    warnings: list[str] = []
    all_papers: list[dict[str, Any]] = []
    for query_index, query in enumerate(queries):
        try:
            all_papers.extend(search_semantic_scholar(query, semantic_key, per_query_limit))
        except Exception as exc:
            warnings.append(f"Semantic Scholar failed for '{query}': {exc}")
        try:
            all_papers.extend(search_openalex(query, per_query_limit))
        except Exception as exc:
            warnings.append(f"OpenAlex failed for '{query}': {exc}")
        if serpapi_key:
            try:
                all_papers.extend(search_serpapi_google_scholar(query, serpapi_key, per_query_limit))
            except Exception as exc:
                warnings.append(f"SerpAPI Scholar failed for '{query}': {exc}")
            try:
                all_papers.extend(search_serpapi_review_layer(query, serpapi_key, max(3, per_query_limit // 2)))
            except Exception as exc:
                warnings.append(f"Review layer failed for '{query}': {exc}")
            if query_index < 3:
                try:
                    all_papers.extend(search_krishikosh_layer(query, serpapi_key, max(3, per_query_limit // 2)))
                except Exception as exc:
                    warnings.append(f"KrishiKosh layer failed for '{query}': {exc}")
        if core_key:
            try:
                all_papers.extend(search_core(query, core_key, per_query_limit))
            except Exception as exc:
                warnings.append(f"CORE failed for '{query}': {exc}")

    deduped = deduplicate_papers(all_papers)
    for paper in deduped:
        score, reason = heuristic_score_paper(paper, context_text)
        paper["score"] = score
        paper["score_reason"] = reason
        paper["category"] = paper.get("category") or infer_paper_category(paper)

    ranked = sorted(deduped, key=lambda item: item.get("score", 0), reverse=True)
    if use_ai_scoring and openai_key:
        ranked = ai_score_papers(openai_key, model, ranked, context_text)
        ranked = sorted(ranked, key=lambda item: item.get("score", 0), reverse=True)

    for paper in ranked:
        paper["selected"] = False
    for paper in ranked[:reference_count]:
        paper["selected"] = True

    return {
        "queries": queries,
        "papers": ranked,
        "selected": ranked[:reference_count],
        "warnings": warnings,
        "candidate_count": len(all_papers),
        "deduped_count": len(deduped),
    }


def pdf_request_headers() -> dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/pdf,application/x-download,text/html,application/xhtml+xml",
        "Referer": "https://scholar.google.com/",
    }


def looks_like_pdf(content: bytes, content_type: str = "") -> bool:
    return bool(content and (content.startswith(b"%PDF") or "pdf" in content_type.lower()))


def download_pdf_url(url: str, timeout: int = 30) -> bytes | None:
    if not url:
        return None
    try:
        response = requests.get(
            url,
            headers=pdf_request_headers(),
            timeout=timeout,
            allow_redirects=True,
            verify=SSL_VERIFY,
        )
        content_type = response.headers.get("Content-Type", "")
        if response.status_code == 200 and looks_like_pdf(response.content, content_type) and len(response.content) > 2000:
            return response.content
    except Exception:
        return None
    return None


def scrape_pdf_links_from_page(url: str, limit: int = 5) -> list[str]:
    if not url or url.lower().endswith(".pdf"):
        return []
    try:
        response = requests.get(url, headers=pdf_request_headers(), timeout=20, verify=SSL_VERIFY)
        if response.status_code != 200:
            return []
        matches = re.findall(r'href=["\']([^"\']+\.pdf(?:\?[^"\']*)?)["\']', response.text, flags=re.IGNORECASE)
        links = []
        for match in matches:
            full_url = urljoin(url, match)
            if full_url not in links:
                links.append(full_url)
        return links[:limit]
    except Exception:
        return []


def extract_doi_from_paper(paper: dict[str, Any]) -> str:
    external_ids = paper.get("external_ids") or {}
    doi = external_ids.get("DOI") or external_ids.get("doi") or ""
    if doi:
        return str(doi).replace("https://doi.org/", "").strip()
    for value in [paper.get("url", ""), paper.get("pdf_url", "")]:
        match = re.search(r"10\.\d{4,9}/[^\s\"<>]+", str(value))
        if match:
            return match.group(0).rstrip(".")
    return ""


def semantic_open_access_pdf_by_title(paper: dict[str, Any], semantic_key: str = "") -> str:
    title = paper.get("title") or ""
    if not title:
        return ""
    try:
        results = search_semantic_scholar(title, semantic_key, limit=1)
        if results:
            return results[0].get("pdf_url") or ""
    except Exception:
        return ""
    return ""


def strategy_direct_pdf(paper: dict[str, Any]) -> tuple[bytes | None, str]:
    candidates = []
    for key in ["pdf_url", "url"]:
        if paper.get(key):
            candidates.append(paper[key])
    candidates.extend(paper.get("pdf_urls") or [])
    for url in candidates:
        pdf = download_pdf_url(str(url))
        if pdf:
            return pdf, "Direct PDF / indexed PDF link"
    link = str(paper.get("url") or "")
    if "arxiv.org" in link:
        arxiv_id = link.rstrip("/").split("/")[-1]
        pdf = download_pdf_url(f"https://arxiv.org/pdf/{arxiv_id}.pdf")
        if pdf:
            return pdf, "arXiv PDF"
    for pdf_link in scrape_pdf_links_from_page(link):
        pdf = download_pdf_url(pdf_link)
        if pdf:
            return pdf, "Scraped open PDF link"
    return None, ""


def strategy_krishikosh_pdf(paper: dict[str, Any]) -> tuple[bytes | None, str]:
    link = str(paper.get("url") or "")
    category = str(paper.get("category") or "")
    if "krishikosh" not in link.lower() and "thesis" not in category.lower():
        return None, ""
    try:
        if link.lower().endswith(".pdf"):
            pdf = download_pdf_url(link)
            if pdf:
                return pdf, "KrishiKosh direct PDF"
        response = requests.get(link, headers=pdf_request_headers(), timeout=20, verify=SSL_VERIFY)
        matches = re.findall(r'href=["\']([^"\']*/bitstream/[^"\']+\.pdf[^"\']*)["\']', response.text, flags=re.IGNORECASE)
        for match in matches:
            full_url = urljoin(link, match)
            pdf = download_pdf_url(full_url)
            if pdf:
                return pdf, "KrishiKosh bitstream"
        if "/handle/" in link:
            handle_id = link.split("/handle/")[-1].strip("/")
            for index in range(1, 5):
                guess_url = f"https://krishikosh.egranth.ac.in/bitstream/1/{handle_id}/{index}/thesis.pdf"
                pdf = download_pdf_url(guess_url)
                if pdf:
                    return pdf, "KrishiKosh guessed bitstream"
    except Exception:
        return None, ""
    return None, ""


def strategy_serpapi_deep_pdf(paper: dict[str, Any], serpapi_key: str = "") -> tuple[bytes | None, str]:
    if not serpapi_key or not paper.get("title"):
        return None, ""
    try:
        response = requests.get(
            "https://serpapi.com/search.json",
            params={"engine": "google_scholar", "q": paper["title"], "api_key": serpapi_key, "num": 1},
            timeout=25,
            verify=SSL_VERIFY,
        )
        response.raise_for_status()
        data = response.json()
        organic = data.get("organic_results", [])
        if not organic:
            return None, ""
        primary = organic[0]
        for resource in primary.get("resources", []) or []:
            if resource.get("file_format") == "PDF" and resource.get("link"):
                pdf = download_pdf_url(resource["link"])
                if pdf:
                    return pdf, "SerpAPI Scholar PDF resource"
        cluster_id = (primary.get("publication_info") or {}).get("cites_id") or primary.get("cluster_id")
        if cluster_id:
            cluster_response = requests.get(
                "https://serpapi.com/search.json",
                params={"engine": "google_scholar", "cluster": cluster_id, "api_key": serpapi_key, "num": 5},
                timeout=25,
                verify=SSL_VERIFY,
            )
            cluster_response.raise_for_status()
            cluster_data = cluster_response.json()
            for result in cluster_data.get("organic_results", []) or []:
                for resource in result.get("resources", []) or []:
                    if resource.get("file_format") == "PDF" and resource.get("link"):
                        pdf = download_pdf_url(resource["link"])
                        if pdf:
                            return pdf, "SerpAPI all-versions PDF"
    except Exception:
        return None, ""
    return None, ""


def strategy_core_pdf(paper: dict[str, Any], core_key: str = "") -> tuple[bytes | None, str]:
    if not core_key:
        return None, ""
    title = paper.get("title") or ""
    if not title:
        return None, ""
    try:
        results = search_core(title, core_key, limit=1)
        if results and results[0].get("pdf_url"):
            pdf = download_pdf_url(results[0]["pdf_url"])
            if pdf:
                return pdf, "CORE API"
    except Exception:
        return None, ""
    return None, ""


def strategy_unpaywall_pdf(paper: dict[str, Any]) -> tuple[bytes | None, str]:
    doi = extract_doi_from_paper(paper)
    if not doi:
        return None, ""
    try:
        response = requests.get(
            f"https://api.unpaywall.org/v2/{doi}",
            params={"email": "research@example.com"},
            timeout=20,
            verify=SSL_VERIFY,
        )
        if response.status_code != 200:
            return None, ""
        data = response.json()
        locations = [data.get("best_oa_location") or {}] + (data.get("oa_locations") or [])
        for location in locations:
            pdf_url = location.get("url_for_pdf")
            if pdf_url:
                pdf = download_pdf_url(pdf_url)
                if pdf:
                    return pdf, "Unpaywall"
    except Exception:
        return None, ""
    return None, ""


def extract_pdf_text_from_bytes(pdf_bytes: bytes) -> str:
    if not pdf_bytes:
        return ""
    try:
        return extract_text_from_pdf(io.BytesIO(pdf_bytes))
    except Exception:
        return ""


def save_reference_pdf(paper: dict[str, Any], pdf_bytes: bytes) -> Path:
    ensure_app_dirs()
    paper_id = safe_slug(str(paper.get("paper_id") or normalize_title(paper.get("title", ""))[:60]))
    title_slug = safe_slug(normalize_title(paper.get("title", ""))[:80])
    filename = f"{paper_id}__{title_slug}.pdf"
    target = DOWNLOADED_REFERENCES_DIR / filename
    target.write_bytes(pdf_bytes)
    return target


def download_and_read_selected_papers(
    selected_papers: list[dict[str, Any]],
    serpapi_key: str = "",
    core_key: str = "",
    semantic_key: str = "",
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for paper in selected_papers:
        enriched = dict(paper)
        pdf_bytes = None
        method = ""

        pdf_url = enriched.get("pdf_url") or semantic_open_access_pdf_by_title(enriched, semantic_key)
        if pdf_url and not enriched.get("pdf_url"):
            enriched["pdf_url"] = pdf_url

        for strategy in [
            strategy_direct_pdf,
            strategy_krishikosh_pdf,
            lambda item: strategy_serpapi_deep_pdf(item, serpapi_key),
            lambda item: strategy_core_pdf(item, core_key),
            strategy_unpaywall_pdf,
        ]:
            pdf_bytes, method = strategy(enriched)
            if pdf_bytes:
                break

        if pdf_bytes:
            pdf_path = save_reference_pdf(enriched, pdf_bytes)
            full_text = extract_pdf_text_from_bytes(pdf_bytes)
            enriched.update(
                {
                    "download_success": True,
                    "download_method": method,
                    "pdf_path": str(pdf_path),
                    "pdf_size_bytes": len(pdf_bytes),
                    "full_text": full_text,
                    "full_text_chars": len(full_text),
                }
            )
        else:
            enriched.update(
                {
                    "download_success": False,
                    "download_method": "",
                    "pdf_path": "",
                    "pdf_size_bytes": 0,
                    "full_text": "",
                    "full_text_chars": 0,
                }
            )
        results.append(enriched)
    return results


def author_surname(author: str) -> str:
    author = (author or "").strip()
    if not author:
        return "Anonymous"
    if "," in author:
        return author.split(",", 1)[0].strip()
    return author.split()[-1].strip()


def citation_key(paper: dict[str, Any]) -> str:
    authors = paper.get("authors") or []
    year = paper.get("year") or "n.d."
    if not authors:
        return f"(Anonymous, {year})"
    first = author_surname(str(authors[0]))
    if len(authors) == 1:
        return f"({first}, {year})"
    return f"({first} et al., {year})"


def format_author_apa(author: str) -> str:
    author = re.sub(r"\s+", " ", author or "").strip()
    if not author:
        return ""
    if "," in author:
        parts = [part.strip() for part in author.split(",", 1)]
        surname = parts[0]
        given = parts[1] if len(parts) > 1 else ""
    else:
        bits = author.split()
        surname = bits[-1]
        given = " ".join(bits[:-1])
    initials = " ".join(f"{part[0].upper()}." for part in re.split(r"\s+|-", given) if part)
    return f"{surname}, {initials}".strip().rstrip(",")


def join_apa_authors(authors: list[str]) -> str:
    formatted = [format_author_apa(author) for author in authors if format_author_apa(author)]
    if not formatted:
        return "Anonymous"
    if len(formatted) == 1:
        return formatted[0]
    if len(formatted) <= 20:
        return ", ".join(formatted[:-1]) + ", & " + formatted[-1]
    return ", ".join(formatted[:19]) + ", ... " + formatted[-1]


def format_apa_reference(paper: dict[str, Any]) -> str:
    authors = join_apa_authors(paper.get("authors") or [])
    year = paper.get("year") or "n.d."
    title = (paper.get("title") or "Untitled").rstrip(".")
    venue = (paper.get("venue") or "").strip()
    url = (paper.get("url") or "").strip()
    doi = (paper.get("external_ids") or {}).get("DOI")
    tail = doi or url
    reference = f"{authors} ({year}). {title}."
    if venue:
        reference += f" {venue}."
    if tail:
        reference += f" {tail}"
    return re.sub(r"\s+", " ", reference).strip()


def reference_context(papers: list[dict[str, Any]], limit: int = 16) -> str:
    rows = []
    for paper in papers[:limit]:
        full_text = paper.get("full_text") or ""
        rows.append(
            {
                "citation": citation_key(paper),
                "title": paper.get("title"),
                "authors": paper.get("authors"),
                "year": paper.get("year"),
                "abstract": truncate_text(paper.get("abstract", ""), 800),
                "full_text_excerpt": truncate_text(full_text, 1800),
                "download_method": paper.get("download_method", ""),
                "full_text_chars": paper.get("full_text_chars", 0),
                "category": paper.get("category") or infer_paper_category(paper),
                "score": paper.get("score"),
                "reason": paper.get("score_reason"),
            }
        )
    return json.dumps(rows, ensure_ascii=True, indent=2)


def section_prompt_common(
    paper_title: str,
    authors: str,
    affiliation: str,
    research_area: str,
    master_context: str,
    analysis: dict[str, Any],
    result_text: str,
) -> str:
    return f"""
Paper title: {paper_title}
Authors/scientists: {authors}
Affiliation/notes: {affiliation}
Research area: {research_area}
Objective: {analysis.get("objective", "")}
Keywords: {analysis.get("keywords", [])}
Major findings: {analysis.get("major_findings", [])}
Treatment rankings: {analysis.get("treatment_rankings", [])}

Master context:
{truncate_text(master_context, 5000)}

Result evidence:
{truncate_text(result_text, 8000)}
"""


def write_methodology(
    api_key: str,
    model: str,
    common: str,
    raw_methodology: str,
    styles: dict[str, str],
) -> str:
    prompt = f"""
Write the Materials and Methods section.

Style source for methodology:
{style_excerpt(styles, "methodology", 6000)}

Required approach:
- Rewrite the raw methodology in passive, factual scientific style.
- Include design, location, season/year, treatments, replications, observations, sampling, and statistics only when supplied.
- Do not invent doses, design, dates, instruments, or statistical tests.
- Use compact agricultural research-paper paragraphs.

{common}

Raw methodology to rewrite:
{truncate_text(raw_methodology, 10000)}
"""
    return chat_text(api_key, model, "You write precise Materials and Methods sections.", prompt, temperature=0.25)


def write_results(
    api_key: str,
    model: str,
    common: str,
    styles: dict[str, str],
    table_summaries: str,
    image_summaries: str,
) -> str:
    prompt = f"""
Write the Results section only.

Preferred result vocabulary:
{style_excerpt(styles, "word_bank", 5000)}

Rules:
- Use only supplied result evidence.
- Use phrases such as significantly, non-significant, minimum, maximum, highest, least, pooled data only when supported.
- Refer to tables and figures where appropriate.
- Do not compare with other authors in this section.
- If statistical significance is not supplied, describe numerical trends cautiously.

{common}

Table evidence:
{truncate_text(table_summaries, 8000)}

Graph or image evidence:
{truncate_text(image_summaries, 5000)}
"""
    return chat_text(api_key, model, "You write Results sections from supplied data without inventing values.", prompt, temperature=0.2)


def write_discussion(
    api_key: str,
    model: str,
    common: str,
    styles: dict[str, str],
    selected_papers: list[dict[str, Any]],
    results_section: str,
) -> str:
    prompt = f"""
Write the Discussion section.

Discussion style examples:
{style_excerpt(styles, "discussion", 7000)}

Preferred comparison vocabulary:
{style_excerpt(styles, "word_bank", 3500)}

Selected references allowed for citation:
{reference_context(selected_papers)}

Rules:
- Interpret the supplied results and compare them with selected references only.
- Prefer downloaded full_text_excerpt evidence where available; use abstracts only when full text was not downloaded.
- Use APA in-text citations exactly like the provided citation strings.
- Use phrases such as "These findings are in agreement with..." only when a selected reference supports it.
- Do not invent author names, years, or references.
- If references are limited, write a cautious interpretation without fake citations.

{common}

Results section already drafted:
{truncate_text(results_section, 7000)}
"""
    return chat_text(api_key, model, "You write discussion sections with careful literature comparison.", prompt, temperature=0.25)


def write_introduction(
    api_key: str,
    model: str,
    common: str,
    styles: dict[str, str],
    selected_papers: list[dict[str, Any]],
) -> str:
    prompt = f"""
Write the Introduction section.

Introduction style examples:
{style_excerpt(styles, "introduction", 5500)}

Dictionary and transition guidance:
{style_excerpt(styles, "introduction_dictionary", 4500)}

Selected references allowed for citation:
{reference_context(selected_papers, limit=10)}

Rules:
- Start broad with crop/host importance when available.
- Explain pest/pathogen/problem severity.
- Identify the knowledge gap or need for the study.
- End with a clear objective sentence.
- Cite only selected references. Do not invent citations.

{common}
"""
    return chat_text(api_key, model, "You write agricultural research-paper introductions.", prompt, temperature=0.3)


def write_conclusion(
    api_key: str,
    model: str,
    common: str,
    styles: dict[str, str],
    results_section: str,
    discussion_section: str,
) -> str:
    prompt = f"""
Write the Conclusion section.

Preferred wording and evaluation phrases:
{style_excerpt(styles, "word_bank", 5000)}

Rules:
- Summarize only the main supported findings.
- Identify the best treatment, factor, relationship, or practical implication when supplied.
- Keep it concise and publication-ready.
- Do not add new data or new citations.

{common}

Results:
{truncate_text(results_section, 5000)}

Discussion:
{truncate_text(discussion_section, 5000)}
"""
    return chat_text(api_key, model, "You write concise scientific conclusions.", prompt, temperature=0.25)


def write_abstract(
    api_key: str,
    model: str,
    common: str,
    styles: dict[str, str],
    methodology_section: str,
    results_section: str,
    conclusion_section: str,
) -> str:
    prompt = f"""
Write the Abstract.

Abstract style examples:
{style_excerpt(styles, "abstract", 6500)}

Rules:
- Follow field-study abstract structure: purpose, place/season if available, method, major results, conclusion.
- Write one compact paragraph unless the target journal clearly requires otherwise.
- Do not cite references in the abstract.
- Do not invent numbers.

{common}

Methodology:
{truncate_text(methodology_section, 4500)}

Results:
{truncate_text(results_section, 4500)}

Conclusion:
{truncate_text(conclusion_section, 2500)}
"""
    return chat_text(api_key, model, "You write concise scientific abstracts.", prompt, temperature=0.25)


def generate_full_draft(
    api_key: str,
    model: str,
    paper_title: str,
    authors: str,
    affiliation: str,
    research_area: str,
    master_context: str,
    raw_methodology: str,
    extracted_files: list[ExtractedFile],
    styles: dict[str, str],
    selected_papers: list[dict[str, Any]],
) -> dict[str, Any]:
    result_text = combined_uploaded_text(extracted_files)
    analysis = analyze_research_context(
        api_key,
        model,
        paper_title,
        research_area,
        master_context,
        raw_methodology,
        result_text,
    )
    title = paper_title.strip() or analysis.get("generated_title") or "Research Paper Draft"
    common = section_prompt_common(title, authors, affiliation, research_area, master_context, analysis, result_text)

    tables = collect_tables(extracted_files)
    images = collect_images(extracted_files)
    table_summaries = "\n\n".join(table_to_plain_text(table) for table in tables)
    image_summaries = "\n\n".join(
        f"Figure from {image['name']}:\n{image.get('summary', '')}" for image in images if image.get("summary")
    )

    methodology = write_methodology(api_key, model, common, raw_methodology, styles)
    results = write_results(api_key, model, common, styles, table_summaries, image_summaries)
    discussion = write_discussion(api_key, model, common, styles, selected_papers, results)
    conclusion = write_conclusion(api_key, model, common, styles, results, discussion)
    introduction = write_introduction(api_key, model, common, styles, selected_papers)
    abstract = write_abstract(api_key, model, common, styles, methodology, results, conclusion)
    references = [format_apa_reference(paper) for paper in selected_papers]

    return {
        "title": title,
        "authors": authors,
        "affiliation": affiliation,
        "abstract": abstract,
        "introduction": introduction,
        "methodology": methodology,
        "results": results,
        "discussion": discussion,
        "conclusion": conclusion,
        "references": references,
        "analysis": analysis,
        "tables": tables,
        "images": images,
        "selected_papers": selected_papers,
    }


def add_paragraphs(document: Any, text: str) -> None:
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text or "") if part.strip()]
    if not paragraphs and text:
        paragraphs = [text.strip()]
    for paragraph_text in paragraphs:
        document.add_paragraph(paragraph_text)


def add_docx_table(document: Any, table: dict[str, Any]) -> None:
    rows = table.get("rows") or []
    if not rows:
        return
    document.add_paragraph(table.get("name") or table.get("source_file") or "Table")
    row_count = min(len(rows), 60)
    col_count = max(len(row) for row in rows[:row_count])
    doc_table = document.add_table(rows=row_count, cols=col_count)
    doc_table.style = "Table Grid"
    for row_index, row in enumerate(rows[:row_count]):
        for col_index in range(col_count):
            doc_table.cell(row_index, col_index).text = str(row[col_index]) if col_index < len(row) else ""
    if len(rows) > row_count:
        document.add_paragraph(f"Note. Table truncated to first {row_count} rows for export.")


def build_docx(draft: dict[str, Any]) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx is required to export DOCX files.")
    document = Document()
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

    section = document.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(draft.get("title", "Research Paper Draft"))
    title_run.bold = True
    title_run.font.size = Pt(14)

    if draft.get("authors"):
        authors = document.add_paragraph()
        authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
        authors.add_run(draft["authors"])
    if draft.get("affiliation"):
        affiliation = document.add_paragraph()
        affiliation.alignment = WD_ALIGN_PARAGRAPH.CENTER
        affiliation.add_run(draft["affiliation"])

    sections = [
        ("Abstract", draft.get("abstract", "")),
        ("Introduction", draft.get("introduction", "")),
        ("Materials and Methods", draft.get("methodology", "")),
        ("Results", draft.get("results", "")),
    ]
    for heading, body in sections:
        document.add_heading(heading, level=1)
        add_paragraphs(document, body)

    for index, table in enumerate(draft.get("tables", []), start=1):
        export_table = dict(table)
        export_table["name"] = export_table.get("name") or f"Table {index}"
        add_docx_table(document, export_table)

    for index, image in enumerate(draft.get("images", []), start=1):
        document.add_paragraph(f"Figure {index}. {image.get('name', 'Uploaded graph')}")
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(image.get("name", "image.png")).suffix) as tmp:
                tmp.write(image["bytes"])
                tmp_path = tmp.name
            document.add_picture(tmp_path, width=Inches(5.8))
            Path(tmp_path).unlink(missing_ok=True)
            if image.get("summary"):
                document.add_paragraph(f"Note. {image['summary']}")
        except Exception as exc:
            document.add_paragraph(f"Image could not be embedded: {exc}")

    document.add_heading("Discussion", level=1)
    add_paragraphs(document, draft.get("discussion", ""))

    document.add_heading("Conclusion", level=1)
    add_paragraphs(document, draft.get("conclusion", ""))

    document.add_heading("References", level=1)
    for reference in draft.get("references", []):
        document.add_paragraph(reference)

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _obj_value(obj: Any, *names: str) -> Any:
    for name in names:
        if isinstance(obj, dict) and name in obj:
            return obj[name]
        if hasattr(obj, name):
            return getattr(obj, name)
    return None


def extract_gemini_grounding(response: Any) -> dict[str, Any]:
    """Return web sources and search queries from a Gemini grounded response."""
    candidates = _obj_value(response, "candidates") or []
    first_candidate = candidates[0] if candidates else None
    grounding = _obj_value(first_candidate, "grounding_metadata", "groundingMetadata") if first_candidate else None
    if not grounding:
        return {"sources": [], "queries": []}

    chunks = _obj_value(grounding, "grounding_chunks", "groundingChunks") or []
    sources: list[dict[str, str]] = []
    seen_urls: set[str] = set()
    for chunk in chunks:
        web = _obj_value(chunk, "web") or {}
        uri = str(_obj_value(web, "uri") or "").strip()
        title = str(_obj_value(web, "title") or "").strip()
        if uri and uri not in seen_urls:
            sources.append({"title": title or uri, "url": uri})
            seen_urls.add(uri)

    queries = []
    for query in _obj_value(grounding, "web_search_queries", "webSearchQueries") or []:
        query_text = str(query).strip()
        if query_text and query_text not in queries:
            queries.append(query_text)

    return {"sources": sources, "queries": queries}


def gemini_generate_text(
    api_key: str,
    model: str,
    prompt: str,
    temperature: float = 0.25,
    use_google_search: bool = True,
    response_mime_type: str | None = None,
) -> dict[str, Any]:
    try:
        from google import genai
        from google.genai import types
    except Exception as exc:  # pragma: no cover - depends on deployment environment
        raise RuntimeError(
            "The google-genai package is required for Gemini article generation. "
            "Install requirements.txt or add google-genai to the Streamlit deployment."
        ) from exc

    client_kwargs: dict[str, Any] = {}
    if api_key:
        client_kwargs["api_key"] = api_key
    client = genai.Client(**client_kwargs)

    config_kwargs: dict[str, Any] = {"temperature": temperature}
    if response_mime_type:
        config_kwargs["response_mime_type"] = response_mime_type
    if use_google_search:
        config_kwargs["tools"] = [types.Tool(google_search=types.GoogleSearch())]
    config = types.GenerateContentConfig(**config_kwargs)

    response = client.models.generate_content(model=model, contents=prompt, config=config)
    text = getattr(response, "text", "") or ""
    grounding = extract_gemini_grounding(response)
    return {"text": text, "sources": grounding["sources"], "queries": grounding["queries"], "raw_response": response}


def month_activity_note(month: str) -> str:
    return GUJARAT_CROP_ACTIVITY_BY_MONTH.get(month, "seasonal crop protection, field monitoring, and timely pest management")


def clean_json_text(text: str) -> str:
    candidate = (text or "").strip()
    candidate = re.sub(r"^```(?:json)?", "", candidate, flags=re.IGNORECASE).strip()
    candidate = re.sub(r"```$", "", candidate).strip()
    return candidate


def topic_discovery_fallback(month: str, year: int, region_scope: str, focus_area: str) -> dict[str, Any]:
    return {
        "research_date": date.today().isoformat(),
        "month": f"{month} {year}",
        "region_scope": region_scope,
        "focus_area": focus_area,
        "topic_candidates": [],
        "selected_topic": {
            "topic": "Seasonal insect and mite monitoring in Gujarat crops",
            "focus_area": focus_area,
            "crop": "",
            "regions": [region_scope],
            "farmer_problem": "Farmers need timely monitoring and integrated management advice for the current season.",
            "why_prevailing": "",
            "month_relevance": month_activity_note(month),
            "recommended_article_angle": "Preventive crop-health monitoring and need-based integrated management.",
            "management_hooks": ["regular scouting", "economic threshold-based decisions", "natural enemy conservation"],
            "score": 0,
        },
        "search_notes": ["Gemini response could not be parsed as JSON; using a safe seasonal fallback."],
    }


def discover_gujarat_agro_topics(
    api_key: str,
    model: str,
    month: str,
    year: int,
    region_scope: str,
    focus_area: str,
    crop_filter: str = "",
    extra_context: str = "",
    max_topics: int = 6,
) -> dict[str, Any]:
    fallback = topic_discovery_fallback(month, year, region_scope, focus_area)
    prompt = f"""
You are a senior agricultural extension researcher preparing topic options for Agro Sandesh.
Use Google Search grounding to research current and prevailing agriculture topics.

Today: {date.today().isoformat()}
Target month: {month} {year}
Region scope: {region_scope}
Focus discipline: {focus_area}
Crop or commodity preference: {crop_filter or "No fixed crop; choose by farmer relevance"}
Seasonal agriculture activity clue for Gujarat: {month_activity_note(month)}
Additional local notes from user:
{truncate_text(extra_context, 2500)}

Research task:
- Find timely and prevailing topics in South Gujarat and wider Gujarat.
- Keep the focus on agricultural acarology, agricultural entomology, or their practical overlap.
- Prioritize current seasonal advisories, pest alerts, agromet advisories, KVK/SAU/ICAR/state agriculture information,
  and credible agriculture news when official sources are not enough.
- Consider crop stage, month relevance, farmer losses, actionability, scientific reliability, and Agro Sandesh fit.
- Do not invent outbreaks or advisories. If evidence is weak, say evidence is limited.

Return valid JSON only with this schema:
{{
  "research_date": "YYYY-MM-DD",
  "month": "{month} {year}",
  "region_scope": "{region_scope}",
  "focus_area": "{focus_area}",
  "topic_candidates": [
    {{
      "topic": "specific article topic",
      "focus_area": "Agricultural entomology or Agricultural acarology or Both",
      "crop": "crop/commodity",
      "regions": ["district/region names"],
      "farmer_problem": "field problem farmers observe",
      "why_prevailing": "why this is relevant now",
      "month_relevance": "relation to {month} activities",
      "evidence_signals": ["short source-backed signals"],
      "recommended_article_angle": "farmer-facing article angle",
      "management_hooks": ["practical management hooks"],
      "score": 0
    }}
  ],
  "selected_topic": {{ same keys as one best candidate }},
  "selection_reason": "why this topic is the best for Agro Sandesh now",
  "search_notes": ["limitations, uncertainty, and useful source notes"]
}}

Return {max_topics} candidates. Score from 0 to 100. Select exactly one best topic.
"""
    result = gemini_generate_text(
        api_key=api_key,
        model=model,
        prompt=prompt,
        temperature=0.15,
        use_google_search=True,
        response_mime_type="application/json",
    )
    parsed = parse_json_object(clean_json_text(result["text"]), fallback)
    if not isinstance(parsed.get("selected_topic"), dict):
        candidates = parsed.get("topic_candidates") or []
        parsed["selected_topic"] = candidates[0] if candidates else fallback["selected_topic"]
    parsed.setdefault("research_date", date.today().isoformat())
    parsed.setdefault("month", f"{month} {year}")
    parsed.setdefault("region_scope", region_scope)
    parsed.setdefault("focus_area", focus_area)
    parsed.setdefault("topic_candidates", [])
    parsed.setdefault("search_notes", [])
    parsed["sources"] = result["sources"]
    parsed["search_queries"] = result["queries"]
    parsed["raw_research_text"] = result["text"]
    return parsed


def generate_agro_sandesh_article(
    api_key: str,
    model: str,
    topic_context: dict[str, Any],
    month: str,
    year: int,
    region_scope: str,
    focus_area: str,
    article_length: str = "1000-1500 words",
    local_observations: str = "",
    extra_requirements: str = "",
) -> dict[str, Any]:
    selected_topic = topic_context.get("selected_topic") or {}
    research_context = json.dumps(
        {
            "topic_candidates": topic_context.get("topic_candidates", []),
            "selection_reason": topic_context.get("selection_reason", ""),
            "search_notes": topic_context.get("search_notes", []),
            "sources": topic_context.get("sources", []),
        },
        ensure_ascii=False,
        indent=2,
    )
    prompt = f"""
You are writing a Gujarati agricultural extension article for Agro Sandesh.
Write in a farmer-centric, scientifically accurate public-extension style inspired by
Dr. M. S. Swaminathan's farmer-welfare communication, without claiming to be him.

Article language: Gujarati
Article length: {article_length}
Target month: {month} {year}
Region scope: {region_scope}
Focus discipline: {focus_area}
Seasonal activity clue: {month_activity_note(month)}

Selected topic context:
{json.dumps(selected_topic, ensure_ascii=False, indent=2)}

Other topic candidates and research notes:
{research_context}

Local field observations or user notes:
{truncate_text(local_observations, 3500)}

Additional article requirements:
{truncate_text(extra_requirements, 2500)}

Writing guide:
{AGRO_SANDESH_STYLE_GUIDE}

Important rules:
- Use Gujarati script.
- Start with a real farmer field situation, not a scientific definition.
- Explain every technical term immediately in simple Gujarati.
- Give step-by-step recommendations. For each recommendation, explain what to do, why it matters, and farmer benefit.
- Mention Gujarat and South Gujarat conditions when supported by the researched context.
- Keep the tone practical, trustworthy, hopeful, and solution-oriented.
- Do not write as a research paper.
- Do not invent pesticide doses, waiting periods, outbreak claims, or district-specific facts not supported by context.
- If a pesticide or biocontrol suggestion needs local label verification, tell farmers to confirm with the local KVK,
  agriculture university, or agriculture department before application.
- End with a practical takeaway message that motivates timely farmer action.

Return valid JSON only:
{{
  "title_gujarati": "article title",
  "topic_english": "selected topic in English",
  "article_gujarati": "full article with farmer-friendly section headings",
  "practical_takeaway_gujarati": "short final takeaway",
  "final_checklist": {{
    "farmer_oriented_opening": true,
    "science_simplified": true,
    "actionable_recommendations": true,
    "hope_and_solutions": true,
    "magazine_not_research_paper": true
  }},
  "source_use_note": "how sources were used and any limitations"
}}
"""
    result = gemini_generate_text(
        api_key=api_key,
        model=model,
        prompt=prompt,
        temperature=0.35,
        use_google_search=True,
        response_mime_type="application/json",
    )
    fallback = {
        "title_gujarati": str(selected_topic.get("topic") or "Agro Sandesh Article"),
        "topic_english": str(selected_topic.get("topic") or ""),
        "article_gujarati": result["text"],
        "practical_takeaway_gujarati": "",
        "final_checklist": {},
        "source_use_note": "Gemini response could not be parsed as JSON; article text is shown as returned.",
    }
    parsed = parse_json_object(clean_json_text(result["text"]), fallback)
    parsed["sources"] = result["sources"] or topic_context.get("sources", [])
    parsed["search_queries"] = result["queries"]
    parsed["topic_context"] = topic_context
    return parsed


def build_agro_sandesh_docx(article: dict[str, Any]) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx is required to export DOCX files.")
    document = Document()
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Nirmala UI"
    normal_style.font.size = Pt(12)

    section = document.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(article.get("title_gujarati") or "Agro Sandesh Article")
    title_run.bold = True
    title_run.font.size = Pt(16)

    if article.get("topic_english"):
        topic = document.add_paragraph()
        topic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        topic.add_run(str(article["topic_english"]))

    add_paragraphs(document, article.get("article_gujarati", ""))

    if article.get("practical_takeaway_gujarati"):
        document.add_heading("Practical takeaway", level=1)
        add_paragraphs(document, article["practical_takeaway_gujarati"])

    sources = article.get("sources") or []
    if sources:
        document.add_heading("Sources checked", level=1)
        for source in sources:
            document.add_paragraph(f"{source.get('title', 'Source')}: {source.get('url', '')}")

    if article.get("source_use_note"):
        document.add_heading("Source note", level=1)
        add_paragraphs(document, article["source_use_note"])

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()
