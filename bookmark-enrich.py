import argparse
import time
from pathlib import Path
import requests
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
import hashlib
import os
import re
from collections import Counter, defaultdict
import json
from datetime import datetime, timezone
from google import genai
import unicodedata
from io import BytesIO
from PIL import Image


DEFAULT_TIMEOUT = 10  # seconds
REQUEST_DELAY = 1.0   # polite delay between requests

INCLUDE_FIELDS = [
    "Title",
    "URL",
    "Description",
    "Favicon Filename",
    "Tags",
    "Priority",
]

GENERIC_TITLE_KEYWORDS = {
    "home", "welcome", "index", "homepage", "untitled", "landing", "start", "default", "intro", "about", "main",
    "default page", "test", "coming soon", "not found"
}

STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "are", "was", "were",
    "into", "about", "your", "their", "they", "them", "you", "its", "using",
    "use", "used", "based", "online", "platform", "service", "tool"
}

TOP_N_KEYWORDS = 8

EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001F5FF"
    "\U0001F600-\U0001F64F"
    "\U0001F680-\U0001F6FF"
    "\U0001F700-\U0001F77F"
    "\U0001F780-\U0001F7FF"
    "\U0001F800-\U0001F8FF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA00-\U0001FAFF"
    "\U00002700-\U000027BF"
    "]+",
    flags=re.UNICODE
)

COMMON_TLDS = [
    ".com", ".net", ".org", ".io", ".ai", ".app", ".dev", ".co",
    ".co.uk", ".me", ".gg", ".tv", ".info", ".xyz"
]

DESCRIPTION_PROMPT_TEMPLATE = """
You are an assistant that writes short, factual descriptions of websites and online tools.

TASK:
1. Write a concise description of the website using information from the provided inputs.
2. After the description, provide a confidence rating (1–5) indicating how accurate the description is.

INPUTS:
- URL: {url}
- Title: {title}
- Metadata Description (may be empty, "Not Available", non-English, or undecipherable): {metadata_description}

DESCRIPTION RULES:
- Write 1 concise sentence (a 2nd very short sentence only if strictly needed).
- If a metadata description is provided AND is not "Not Available":
  - Use it as the primary factual source.
  - If it is non-English, translate it to English before reasoning.
  - Prefer concrete nouns, functions, and actions.
  - Exclude slogans, hype, and vague marketing claims.
- If metadata is empty, "Not Available" or undecipherable, rely solely on the URL and page title.
  - If URL and page title does not provide sufficient information to construct a description, output the description as: "Information insufficient".
- Start with a noun phrase naming the resource type (e.g., "Web-based tool", "Documentation site", "Open-source library").
- Focus on what the site does, not opinions or quality.
- Avoid marketing language, promotional adjectives, or vague claims.
- Use plain text only (no emojis, lists, or quotes).
- Description length: 12–23 words.
- Follow this pattern:
  "[Noun phrase] for/that + [brief function or capability] + [optional context or audience]."

CONFIDENCE RATING RULES:
- Provide a confidence rating (1–5) indicating how well the description is supported by verifiable information, particularly alignment with the metadata description and title.
- Use the Confidence Rating Scale as a reference for your confidence rating.
- For each of your confidence rating, provide a short explanation of why you chose that rating.
    - If there are extra notes about the confidence rating, include them inline within the Explanation text, without adding new section labels.
    - Do not introduce additional sections or labels.
    - Do not repeat or restate the confidence rating scale.
- If Description is Information insufficient, return confidence = 0.
- If Metadata Description is empty, "Not Available" or undecipherable, start with a confidence of 1.
- Otherwise, start with a base confidence of 2.
- Use the following criteria to adjust confidence rating:
  - Overlap between nouns and functional phrases in the final description and the metadata.
  - Overlap between nouns and functional phrases in the final description and the title.
  - Whether the description avoids marketing language present in the metadata.
  - Whether the description stays factual and specific.
- Penalize:
  - Added features not supported by metadata
  - Generic or inferred functionality
  - Marketing language (e.g., "best", "powerful", "ultimate")

Confidence Rating Scale:
0 = Information insufficient (Reserved for no description available)
1 = Very low confidence (guessing, insufficient or unclear information)
2 = Low confidence (some clues but weak or ambiguous)
3 = Moderate confidence (general purpose inferred, limited specificity)
4 = High confidence (clear alignment with metadata and title)
5 = Very high confidence (description directly supported by metadata description)

OUTPUT FORMAT:
<description> [Confidence: X] Explanation: Y

Examples:
Multi-scanner file analysis platform with threat scoring and metadata inspection. [Confidence: 5] Explanation: Explicit, precise match to metadata description
Interactive web quiz for practicing Japanese verb conjugations with timed exercises. [Confidence: 4] Explanation: Clear alignment with metadata and title

Site information:
URL: {url}
Title: {title}
Metadata Description: {metadata_description}

Write the description now.

""".strip()

GEMINI_MODEL = "gemini-2.5-flash"


def get_genai_client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY environment variable is not set")
    return genai.Client(api_key=api_key)


def load_workbook_safe(path):
    wb = load_workbook(path, keep_vba=True)
    if "Bookmarks" not in wb.sheetnames:
        raise ValueError("Sheet 'Bookmarks' not found")
    return wb, wb["Bookmarks"]


def ensure_column(ws, column_name):
    headers = [cell.value for cell in ws[1]]
    if column_name in headers:
        return headers.index(column_name) + 1

    ws.cell(row=1, column=len(headers) + 1, value=column_name)
    return len(headers) + 1


def check_url_status(url):
    try:
        response = requests.head(
            url,
            allow_redirects=True,
            timeout=DEFAULT_TIMEOUT,
            headers={"User-Agent": "BookmarkStatusChecker/1.0"},
        )
        return f"{response.status_code}"
    except requests.exceptions.Timeout:
        return "timeout"
    except requests.exceptions.ConnectionError:
        return "connection_error"
    except requests.exceptions.RequestException as e:
        return f"error"

def extract_title_from_url(url):
    try:
        response = requests.get(
            url,
            timeout=DEFAULT_TIMEOUT,
            headers={"User-Agent": "BookmarkTitleExtractor/1.0"},
        )
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # 1. Open Graph title
        og_title = soup.find("meta", property="og:title")
        if og_title and og_title.get("content"):
            return og_title["content"].strip()

        # 2. HTML <title>
        if soup.title and soup.title.string:
            return soup.title.string.strip()

    except requests.exceptions.RequestException:
        pass

    # 3. Fallback: domain name
    parsed = urlparse(url)
    return parsed.netloc

def strip_emojis(text: str) -> str:
    if not text:
        return text
    return EMOJI_PATTERN.sub("", text).strip()

def google_favicon_url(url, size=64):
    parsed = urlparse(url)
    domain = parsed.netloc
    if not domain:
        return None
    return f"https://www.google.com/s2/favicons?domain={domain}&sz={size}"


def favicon_exists(url):
    try:
        response = requests.head(
            url,
            timeout=DEFAULT_TIMEOUT,
            headers={"User-Agent": "BookmarkFaviconChecker/1.0"},
        )
        return response.status_code == 200
    except requests.exceptions.RequestException:
        return False


def is_valid_google_favicon(url):
    """
    Check if the Google S2 service returned a valid favicon or the default globe.
    We request sz=64. If the returned image is 16x16, it is the default icon.
    """
    try:
        response = requests.get(
            url,
            timeout=DEFAULT_TIMEOUT,
            headers={"User-Agent": "BookmarkFaviconChecker/1.0"},
        )
        if response.status_code != 200:
            return False

        # Check image dimensions
        try:
            img = Image.open(BytesIO(response.content))
            # Google default icon stays 16x16 even if we requested sz=64
            if img.size == (16, 16):
                return False
            return True
        except Exception:
            # If we can't parse it as an image, assume it's invalid
            return False

    except requests.exceptions.RequestException:
        return False


def parse_favicon_from_html(page_url):
    try:
        response = requests.get(
            page_url,
            timeout=DEFAULT_TIMEOUT,
            headers={"User-Agent": "BookmarkFaviconParser/1.0"},
        )
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        icon_rels = [
            "icon",
            "shortcut icon",
            "apple-touch-icon",
            "apple-touch-icon-precomposed",
        ]

        for rel in icon_rels:
            tag = soup.find("link", rel=lambda x: x and rel in x.lower())
            if tag and tag.get("href"):
                return urljoin(page_url, tag["href"])

    except requests.exceptions.RequestException:
        pass

    return None

def safe_favicon_filename(url, favicon_url):
    """
    Generate a deterministic, filesystem-safe filename.
    """
    parsed = urlparse(url)

    ext = os.path.splitext(favicon_url)[1]
    if not ext or len(ext) > 5:
        ext = ".ico"

    digest = hashlib.sha256(favicon_url.encode("utf-8")).hexdigest()[:8]
    # Use site URL for base name to avoid generic names (e.g. google) and better TLD stripping
    clean_base = clean_domain_for_filename(url)

    return f"{clean_base}_{digest}{ext}"

def clean_domain_for_filename(url: str) -> str:
    """
    Extracts a clean base name from a URL:
    - Removes scheme
    - Removes www.
    - Removes common TLDs
    """

    parsed = urlparse(url)
    host = parsed.netloc or parsed.path
    if host:
        host = host.strip()

    host = host.lower()

    # remove port if present
    host = host.split(":")[0]

    # remove www.
    if host.startswith("www."):
        host = host[4:]

    # remove known TLDs (longest first)
    for tld in sorted(COMMON_TLDS, key=len, reverse=True):
        if host.endswith(tld):
            host = host[: -len(tld)]
            break

    # final safety cleanup
    host = re.sub(r"[^a-z0-9\-]+", "_", host)
    host = host.strip("_")

    return host or "site"

def download_favicon(
    favicon_url: str,
    page_url: str,
    output_dir: str,
) -> str | None:
    """
    Downloads a favicon, normalizes it to PNG,
    and names it based on the page URL (not favicon URL).
    """

    try:
        resp = requests.get(favicon_url, timeout=10)
        if resp.status_code != 200:
            return None

        # --- Normalize image to PNG ---
        image = Image.open(BytesIO(resp.content)).convert("RGBA")

        # --- Stable uniqueness from favicon content ---
        digest = hashlib.sha256(resp.content).hexdigest()[:8]

        # --- Clean base name from PAGE URL ---
        base = clean_domain_for_filename(page_url)

        filename = f"{base}_{digest}.png"
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)

        image.save(path, format="PNG")

        return filename

    except Exception as e:
        print(f"Failed to download favicon for {page_url}: {e}")
        return None

def is_likely_non_english(text, threshold=0.3):
    if not text or not isinstance(text, str):
        return False

    ascii_chars = sum(1 for c in text if ord(c) < 128)
    ratio = ascii_chars / max(len(text), 1)

    return ratio < (1 - threshold)

def extract_metadata_description(url, timeout=10):
    """
    Extract a site's description using common metadata standards.
    Priority:
    1. meta[name="description"]
    2. meta[property="og:description"]
    3. meta[name="twitter:description"]
    4. JSON-LD: description -> headline
    5. <title>
    Returns a string, or "Not Available".
    """
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (compatible; BookmarkBot/1.0)"
        }
        resp = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)

        if resp.status_code != 200 or not resp.text:
            return "Not Available"

        soup = BeautifulSoup(resp.text, "html.parser")

        # 1. <meta name="description">
        tag = soup.find("meta", attrs={"name": "description"})
        if tag and tag.get("content"):
            return tag["content"].strip()

        # 2. <meta property="og:description">
        tag = soup.find("meta", attrs={"property": "og:description"})
        if tag and tag.get("content"):
            return tag["content"].strip()

        # 3. <meta name="twitter:description">
        tag = soup.find("meta", attrs={"name": "twitter:description"})
        if tag and tag.get("content"):
            return tag["content"].strip()

        # 4. JSON-LD
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string)

                # JSON-LD can be a list or dict
                nodes = data if isinstance(data, list) else [data]

                for node in nodes:
                    if isinstance(node, dict):
                        if "description" in node and isinstance(node["description"], str):
                            return node["description"].strip()
                        if "headline" in node and isinstance(node["headline"], str):
                            return node["headline"].strip()
            except Exception:
                continue

        # 5. <title> fallback
        if soup.title and soup.title.string:
            return soup.title.string.strip()

        return "Not Available"

    except Exception as e:
        print(f"Metadata extraction failed: {url} ({e})")
        return "Not Available"

def get_or_create_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    return wb.create_sheet(sheet_name)

def fetch_duckduckgo_description(query, max_length=600):
    params = {
        "q": query,
        "format": "json",
        "no_redirect": 1,
        "no_html": 1,
        "skip_disambig": 1,
    }

    try:
        resp = requests.get(
            "https://api.duckduckgo.com/",
            params=params,
            timeout=15
        )
        resp.raise_for_status()
        data = resp.json()

        parts = []

        # Primary abstract (best source)
        abstract = data.get("AbstractText")
        if abstract:
            parts.append(abstract.strip())

        # Related topics (secondary source)
        for item in data.get("RelatedTopics", []):
            if isinstance(item, dict) and item.get("Text"):
                parts.append(item["Text"].strip())

            if sum(len(p) for p in parts) >= max_length:
                break

        if parts:
            combined = " ".join(parts)
            return combined[:max_length].strip()

    except Exception as e:
        print(f"DuckDuckGo error for query '{query}': {e}")

    return None

def fetch_duckduckgo_html_snippet(query, timeout=15):
    url = "https://duckduckgo.com/html/"
    params = {"q": query}

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    try:
        resp = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=timeout
        )
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")

        # Find first result snippet
        snippet = soup.select_one("a.result__snippet")

        if snippet:
            text = snippet.get_text(strip=True)
            return text if text else None

    except Exception as e:
        print(f"DuckDuckGo HTML scrape error for '{query}': {e}")

    return None

def normalize_text(text):
    if not text:
        return ""

    text = unicodedata.normalize("NFKD", text)
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize_words(text, min_len=1):
    text = normalize_text(text)
    return re.findall(rf"\b[a-z0-9]{{{min_len},}}\b", text)

def extract_domain_tokens(url):
    try:
        netloc = urlparse(url).netloc.replace("www.", "")
        return set(tokenize_words(netloc, min_len=3))
    except Exception:
        return set()


def extract_url_path_tokens(url):
    try:
        path = urlparse(url).path
        return set(tokenize_words(path, min_len=4))
    except Exception:
        return set()


def extract_top_metadata_keywords(metadata):
    words = tokenize_words(metadata, min_len=4)
    words = [w for w in words if w not in STOPWORDS]
    return set(w for w, _ in Counter(words).most_common(TOP_N_KEYWORDS))

def translate_with_ollama(text, model="llama3.2:3b"):
    prompt = (
        "If the text is already in English, return it unchanged.\n"
        "If the following text is not English, translate it to clear, factual English.\n"
        "Do not paraphrase, summarize or rewrite. Do not add information.\n"
        "Return only the translated text.\n\n"
        f"{text}"
    )

    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": model,
                "prompt": prompt,
                "stream": False
            },
            timeout=60
        )
        response.raise_for_status()
        return response.json().get("response", "").strip()
    except Exception as e:
        print(f"Ollama translation error: {e}")
        return None

def generate_description_ollama(url, title, metadata_description, model="llama3.2:3b"):
    clean_metadata = strip_emojis(metadata_description)

    prompt = DESCRIPTION_PROMPT_TEMPLATE.format(
        url=url,
        title=title or "",
        metadata_description=clean_metadata or "Not Available"
    )

    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.2,
            "top_p": 0.9,
        },
    }

    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json=payload,
            timeout=60,
        )
        response.raise_for_status()
        data = response.json()
        return data.get("response", "").strip()

    except requests.exceptions.RequestException as e:
        print(f"Ollama error: {e}")
        return None

def generate_description_gemini(url, title, metadata_description, model=GEMINI_MODEL):
    clean_metadata = strip_emojis(metadata_description)
    prompt = DESCRIPTION_PROMPT_TEMPLATE.format(
        url=url,
        title=title or "",
        metadata_description=clean_metadata or "Not Available"
    )

    try:
        client = get_genai_client()
        response = client.models.generate_content(
            model=model,
            contents=prompt,
            config={
                "temperature": 0.2,
                "top_p": 0.9,
                "max_output_tokens": 80,
            }
        )

        if not response.text:
            return None

        return response.text.strip()

    except Exception as e:
        print(f"Gemini error: {e}")
        return None

def parse_llm_description(response: str):
    """
    Robustly parses:
    - description
    - confidence (int)
    - explanation (str)

    Works even if ordering or newlines vary.
    """

    if not response or not isinstance(response, str):
        return None, None, None

    text = response.strip()

    # --- Extract confidence ---
    confidence = None
    conf_match = re.search(r"\[Confidence:\s*([0-5])\]", text, re.IGNORECASE)
    if conf_match:
        confidence = int(conf_match.group(1))

    # --- Extract explanation ---
    explanation = None
    expl_match = re.search(
        r"Explanation:\s*(.+)",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if expl_match:
        explanation = expl_match.group(1).strip()

    # --- Remove confidence + explanation safely ---
    cleaned = re.sub(r"\[Confidence:\s*[0-5]\]", "", text, flags=re.IGNORECASE)
    cleaned = re.sub(
        r"Explanation:\s*.+",
        "",
        cleaned,
        flags=re.IGNORECASE | re.DOTALL
    )

    description = cleaned.strip()

    if not description:
        description = None

    return description, confidence, explanation


def compute_heuristic_confidence(url, title, description, metadata_description):
    """
    Returns an integer confidence score from 0–5.
    """

    # Rule 2: Information insufficient → hard stop
    if not description or normalize_text(description) == "information insufficient":
        return 0

    score = 3  # Rule 1: neutral start

    # -------------------------
    # Rule 3: Title quality check
    # -------------------------

    if not title or not normalize_text(title):
        score -= 2
    else:
        title_tokens = set(tokenize_words(title))
        if title_tokens & GENERIC_TITLE_KEYWORDS:
            score -= 2

    # -------------------------
    # Rule 4: Domain ↔ title overlap
    # -------------------------

    domain_tokens = extract_domain_tokens(url)
    if title:
        title_tokens = set(tokenize_words(title, min_len=3))
        if domain_tokens & title_tokens:
            score += 1

    # -------------------------
    # Rule 5: Metadata ↔ description overlap
    # -------------------------

    if metadata_description and metadata_description.strip() != "Not Available":
        meta_keywords = extract_top_metadata_keywords(metadata_description)
        desc_keywords = set(tokenize_words(description, min_len=4))

        overlap_count = len(meta_keywords & desc_keywords)

        if overlap_count == 0:
            score -= 2
        elif overlap_count <= 2:
            score += 1
        else:
            score += 2

    # -------------------------
    # Rule 6: URL path ↔ description overlap
    # -------------------------

    path_tokens = extract_url_path_tokens(url)
    desc_tokens = set(tokenize_words(description, min_len=4))

    if path_tokens:
        overlap_count = len(path_tokens & desc_tokens)

        if overlap_count == 0:
            score -= 1
        elif overlap_count <= 2:
            score += 1
        else:
            score += 2

    # -------------------------
    # Rule 7: Clamp final score
    # -------------------------

    return max(1, min(5, score))

def combine_confidence(llm_conf, heuristic_conf):
    """
    Combine LLM self-assessment with heuristics.
    Implements a 'Veto' mechanism: if heuristics are very low,
    we distrust the LLM's high self-confidence (hallucination risk).
    """
    if llm_conf is None:
        return heuristic_conf

    # Veto: If heuristic is 1 (almost certainly bad), cap final at 2
    if heuristic_conf <= 1:
        return min(llm_conf, 2)

    # Soft Veto: If heuristic is 2, cap final at 3
    if heuristic_conf == 2:
        return min(llm_conf, 3)

    combined = round((llm_conf * 0.6) + (heuristic_conf * 0.4))
    return max(1, min(5, combined))

def run_status_check(ws):
    headers = [cell.value for cell in ws[1]]

    if "URL" not in headers:
        raise ValueError("No 'URL' column found")

    url_col = headers.index("URL") + 1
    status_col = ensure_column(ws, "Status")

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        current_status = ws.cell(row=row, column=status_col).value

        if not url:
            continue

        # Skip rows already checked
        if current_status:
            continue

        print(f"Checking: {url}")
        status = check_url_status(url)
        ws.cell(row=row, column=status_col, value=status)

        time.sleep(REQUEST_DELAY)

def run_title_extraction(ws):
    headers = [cell.value for cell in ws[1]]

    if "URL" not in headers:
        raise ValueError("No 'URL' column found")

    url_col = headers.index("URL") + 1
    title_col = ensure_column(ws, "Title")

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        current_title = ws.cell(row=row, column=title_col).value

        if not url:
            continue

        # Skip rows that already have a title
        if current_title:
            continue

        print(f"Extracting title: {url}")
        title = extract_title_from_url(url)
        ws.cell(row=row, column=title_col, value=title)

        time.sleep(REQUEST_DELAY)

def run_favicon_discovery(ws):
    headers = [cell.value for cell in ws[1]]

    if "URL" not in headers:
        raise ValueError("No 'URL' column found")

    url_col = headers.index("URL") + 1
    favicon_col = ensure_column(ws, "Favicon URL")

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        current_favicon = ws.cell(row=row, column=favicon_col).value

        if not url:
            continue

        # Skip rows already processed
        if current_favicon:
            continue

        print(f"Finding favicon: {url}")

        favicon_url = google_favicon_url(url)
        # Check if Google returned a real favicon (not the default globe)
        if favicon_url and is_valid_google_favicon(favicon_url):
            print("  -> Found via Google (primary)")
            ws.cell(row=row, column=favicon_col, value=favicon_url)
            continue

        print("  -> Google returned default/invalid, trying HTML parsing (fallback)...")

        # Fallback to HTML parsing
        parsed_favicon = parse_favicon_from_html(url)
        if parsed_favicon:
            print("  -> Found via HTML parsing")
            ws.cell(row=row, column=favicon_col, value=parsed_favicon)
        else:
             print("  -> No favicon found")

        time.sleep(REQUEST_DELAY)

def run_favicon_download(ws, favicon_dir):
    headers = [cell.value for cell in ws[1]]

    required_cols = ["URL", "Favicon URL"]
    for col in required_cols:
        if col not in headers:
            raise ValueError(f"No '{col}' column found")

    url_col = headers.index("URL") + 1
    favicon_url_col = headers.index("Favicon URL") + 1
    filename_col = ensure_column(ws, "Favicon Filename")

    favicon_dir.mkdir(parents=True, exist_ok=True)

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        favicon_url = ws.cell(row=row, column=favicon_url_col).value
        current_filename = ws.cell(row=row, column=filename_col).value

        if not url or not favicon_url:
            continue

        # Skip already downloaded
        if current_filename:
            continue

        print(f"Downloading favicon: {favicon_url}")

        # New signature: download_favicon(favicon_url, page_url, output_dir)
        # Returns filename on success, None on failure
        saved_filename = download_favicon(favicon_url, url, str(favicon_dir))

        if saved_filename:
            ws.cell(row=row, column=filename_col, value=saved_filename)

        time.sleep(REQUEST_DELAY)

def run_metadata_description_extraction(ws):
    headers = [cell.value for cell in ws[1]]

    url_col = headers.index("URL") + 1
    meta_desc_col = ensure_column(ws, "Metadata Description")

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        existing_meta = ws.cell(row=row, column=meta_desc_col).value

        if not url or existing_meta:
            continue

        print(f"Extracting metadata description: {url}")

        meta_desc = extract_metadata_description(url)

        ws.cell(row=row, column=meta_desc_col, value=meta_desc)

        time.sleep(REQUEST_DELAY)

def normalize_metadata_to_english(wb, source_sheet_name="Bookmarks"):
    ws = wb[source_sheet_name]
    headers = [cell.value for cell in ws[1]]

    if "Metadata Description" not in headers:
        print("No 'Metadata Description' column found.")
        return

    meta_col = headers.index("Metadata Description") + 1

    backup_ws = get_or_create_sheet(wb, "Metadata Backup")

    if backup_ws.max_row == 1:
        backup_ws.append(["Row", "URL", "Original Metadata"])

    url_col = headers.index("URL") + 1 if "URL" in headers else None

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=meta_col)
        metadata = cell.value

        # --- Skip conditions ---
        if not metadata:
            continue

        if not isinstance(metadata, str):
            continue

        metadata_stripped = metadata.strip()

        if metadata_stripped.lower() == "not available":
            continue

        if "🌍" in metadata_stripped:
            continue

        if not is_likely_non_english(metadata_stripped):
            continue

        print(f"Translating metadata (row {row})")

        translated = translate_with_ollama(metadata_stripped)
        if not translated:
            continue

        # Backup original metadata
        backup_ws.append([
            row,
            ws.cell(row=row, column=url_col).value if url_col else "",
            metadata_stripped
        ])

        # Replace with translated text + flag
        cell.value = f"{translated} 🌍"

        time.sleep(1)

def fill_metadata_from_duckduckgo(wb, source_sheet_name="Bookmarks"):
    ws = wb[source_sheet_name]
    headers = [cell.value for cell in ws[1]]

    if "Metadata Description" not in headers:
        print("No 'Metadata Description' column found.")
        return

    meta_col = headers.index("Metadata Description") + 1
    title_col = headers.index("Title") + 1 if "Title" in headers else None
    url_col = headers.index("URL") + 1 if "URL" in headers else None

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=meta_col)
        value = cell.value

        if not value or not isinstance(value, str):
            continue

        if value.strip().lower() != "not available":
            continue

        title = ws.cell(row=row, column=title_col).value if title_col else ""
        url = ws.cell(row=row, column=url_col).value if url_col else ""

        base_query = title or url
        query = f"what is {base_query}"

        if not query:
            continue

        print(f"DuckDuckGo lookup (row {row}): {query}")

        description = fetch_duckduckgo_description(query)
        if not description:
            continue

        cell.value = f"{description} 🦆"

        time.sleep(1)

def run_duckduckgo_html_metadata(ws):
    headers = [cell.value for cell in ws[1]]

    url_col = headers.index("URL") + 1
    title_col = headers.index("Title") + 1
    meta_col = ensure_column(ws, "Metadata Description")

    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        title = ws.cell(row=row, column=title_col).value
        current_meta = ws.cell(row=row, column=meta_col).value

        # Skip if already populated or explicitly unavailable
        if not url or (current_meta and current_meta != "Not Available"):
            continue

        query = title.strip() if isinstance(title, str) and title.strip() else f"what is {url}"

        print(f"DDG HTML scrape: {query}")

        snippet = fetch_duckduckgo_html_snippet(query)

        if snippet:
            ws.cell(
                row=row,
                column=meta_col,
                value=f"{snippet} 🦆"
            )

        # Be polite
        time.sleep(3)

def update_category_overview(wb, source_sheet_name, target_sheet_name="Categories"):
    """
    Count bookmarks per category and write the summary to a separate sheet.
    """
    ws = wb[source_sheet_name]

    headers = [cell.value for cell in ws[1]]

    if "Category" not in headers:
        raise ValueError("No 'Category' column found in source sheet")

    category_col = headers.index("Category") + 1

    counts = Counter()

    for row in range(2, ws.max_row + 1):
        category = ws.cell(row=row, column=category_col).value
        if category:
            counts[category.strip()] += 1

    # Create or clear target sheet
    if target_sheet_name in wb.sheetnames:
        overview_ws = wb[target_sheet_name]
        overview_ws.delete_rows(1, overview_ws.max_row)
    else:
        overview_ws = wb.create_sheet(title=target_sheet_name)

    # Write header
    overview_ws.append(["Category", "Count"])

    # Write sorted results
    for category, count in sorted(counts.items(), key=lambda x: (-x[1], x[0])):
        overview_ws.append([category, count])

def is_private_value(value):
    if value is None:
        return False
    return str(value).strip().lower() in {"yes", "true", "1"}

def export_bookmarks_to_json(
    wb,
    bookmarks_sheet_name="Bookmarks",
    categories_sheet_name="Categories",
    output_path="bookmarks.json",
    include_fields=None,
    exclude_fields=None,
    mode="public",  # "public", "private", "all"
):
    ws = wb[bookmarks_sheet_name]
    cat_ws = wb[categories_sheet_name]

    headers = [cell.value for cell in ws[1]]
    header_index = {h: i + 1 for i, h in enumerate(headers)}

    if include_fields:
        fields = [f for f in include_fields if f in header_index]
    else:
        fields = headers.copy()
        if exclude_fields:
            fields = [f for f in fields if f not in exclude_fields]

    category_counts = {}
    for row in range(2, cat_ws.max_row + 1):
        category = cat_ws.cell(row=row, column=1).value
        count = cat_ws.cell(row=row, column=2).value
        if category:
            category_counts[category] = count

    categories = defaultdict(list)

    total_rows = 0

    private_col = header_index.get("Private?")

    for row in range(2, ws.max_row + 1):
        private_value = None
        if private_col:
            private_value = ws.cell(row=row, column=private_col).value

        # FILTER LOGIC
        if mode == "private":
            # Export ONLY private rows
            if not is_private_value(private_value):
                continue
        elif mode == "public":
            # Default behavior: exclude private rows
            if is_private_value(private_value):
                continue
        # mode == "all": don't filter anything

        category = ws.cell(
            row=row,
            column=header_index.get("Category")
        ).value or "Uncategorized"

        bookmark = {}
        for field in fields:
            bookmark[field] = ws.cell(
                row=row,
                column=header_index[field]
            ).value or ""

        categories[category].append(bookmark)
        total_rows += 1

    output = {
        "last_updated": datetime.now(timezone.utc).isoformat(),
        "total_rows": total_rows,
        "categories": dict(categories)
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"JSON exported to: {output_path}")

def run_description_generation(ws, provider="ollama"):
    headers = [cell.value for cell in ws[1]]

    url_col = headers.index("URL") + 1
    url_col = headers.index("URL") + 1
    title_col = headers.index("Title") + 1

    # Optional: If metadata descriptions exist, use them
    meta_desc_col_idx = None
    if "Metadata Description" in headers:
        meta_desc_col_idx = headers.index("Metadata Description") + 1

    desc_col = ensure_column(ws, "Description")
    llm_conf_col = ensure_column(ws, "LLM Confidence")
    final_conf_col = ensure_column(ws, "Final Confidence")
    llm_expl_col = ensure_column(ws, "LLM Confidence Explanation")


    for row in range(2, ws.max_row + 1):
        url = ws.cell(row=row, column=url_col).value
        title = ws.cell(row=row, column=title_col).value
        existing_desc = ws.cell(row=row, column=desc_col).value

        # Fetch metadata description if available
        metadata_description = None
        if meta_desc_col_idx:
            metadata_description = ws.cell(row=row, column=meta_desc_col_idx).value

        if not url or existing_desc:
            continue

        print(f"Generating description ({provider}): {url}")

        # --- LLM provider selection ---
        if provider == "ollama":
            response = generate_description_ollama(url, title, metadata_description)
        elif provider == "gemini":
            response = generate_description_gemini(url, title, metadata_description)
        else:
            raise ValueError(f"Unknown LLM provider: {provider}")

        description, llm_conf, llm_expl = parse_llm_description(response)

        # -----------------------
        # Insufficient information short-circuit
        # -----------------------
        if description and description.strip().lower().startswith("information insufficient"):
            heuristic_conf = 0
            final_conf = 0
        else:
            heuristic_conf = compute_heuristic_confidence(url, title, description, metadata_description)
            final_conf = combine_confidence(llm_conf, heuristic_conf)

        if description:
            ws.cell(row=row, column=desc_col, value=description)
            ws.cell(row=row, column=llm_conf_col, value=llm_conf)
            ws.cell(row=row, column=llm_expl_col, value=llm_expl)
            ws.cell(row=row, column=final_conf_col, value=final_conf)

        time.sleep(REQUEST_DELAY)

def main():
    parser = argparse.ArgumentParser(description="Bookmark enrichment tool")
    parser.add_argument("workbook", help="Path to bookmarks.xlsm")
    parser.add_argument("--status", action="store_true", help="Check website status")
    parser.add_argument("--titles", action="store_true", help="Extract page titles")
    parser.add_argument("--favicons", action="store_true", help="Discover favicon URLs")
    parser.add_argument(
    "--download-favicons",
    action="store_true",
    help="Download favicon files"
    )
    parser.add_argument(
    "--descriptions",
    action="store_true",
    help="Generate descriptions using LLM"
    )
    parser.add_argument(
    "--llm",
    choices=["ollama", "gemini"],
    default="ollama",
    help="LLM provider for description generation"
    )
    parser.add_argument(
    "--categories",
    action="store_true",
    help="Update category count overview sheet"
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Export bookmarks and categories to JSON"
    )
    parser.add_argument(
    "--metadata",
    action="store_true",
    help="Extract site metadata descriptions into the 'Metadata Description' column"
    )
    parser.add_argument(
        "--normalize-metadata",
        action="store_true",
        help="Translate non-English metadata descriptions to English using Ollama"
    )
    parser.add_argument(
    "--fill-metadata-search",
    action="store_true",
    help="Fill 'Not Available' metadata descriptions using DuckDuckGo search"
    )
    parser.add_argument(
    "--ddg-html",
    action="store_true",
    help="Fill 'Metadata Description' using DuckDuckGo HTML snippet scraping"
    )
    parser.add_argument(
    "--private-only",
    action="store_true",
    help="Export only bookmarks marked as private"
    )
    parser.add_argument(
    "--export-both",
    action="store_true",
    help="Export both public bookmarks (bookmarks.json) and private bookmarks (private.json)"
    )
    parser.add_argument(
        "--output",
        help="Output filename for JSON export (used with --json)"
    )
    args = parser.parse_args()

    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        raise FileNotFoundError(workbook_path)

    wb, ws = load_workbook_safe(workbook_path)

    if args.status:
        run_status_check(ws)

    if args.titles:
        run_title_extraction(ws)

    if args.favicons:
        run_favicon_discovery(ws)

    if args.download_favicons:
        favicon_dir = workbook_path.parent / "favicons"
        run_favicon_download(ws, favicon_dir)

    if args.metadata:
        run_metadata_description_extraction(ws)

    if args.descriptions:
        run_description_generation(ws, provider=args.llm)

    if args.categories:
        update_category_overview(
            wb,
            source_sheet_name=ws.title,
            target_sheet_name="Categories"
        )
    if args.json:
        mode = "private" if args.private_only else "public"
        out_path = args.output or ("private.json" if args.private_only else "bookmarks.json")

        export_bookmarks_to_json(
            wb,
            bookmarks_sheet_name=ws.title,
            categories_sheet_name="Categories",
            output_path=out_path,
            include_fields=INCLUDE_FIELDS,
            mode=mode
        )

    if args.export_both:
        print("Exporting public bookmarks...")
        export_bookmarks_to_json(
            wb,
            bookmarks_sheet_name=ws.title,
            categories_sheet_name="Categories",
            output_path="bookmarks.json",
            include_fields=INCLUDE_FIELDS,
            mode="public"
        )
        print("Exporting private bookmarks...")
        export_bookmarks_to_json(
            wb,
            bookmarks_sheet_name=ws.title,
            categories_sheet_name="Categories",
            output_path="private.json",
            include_fields=INCLUDE_FIELDS,
            mode="private"
        )


    if args.normalize_metadata:
        print("Running metadata normalization...")
        normalize_metadata_to_english(wb, source_sheet_name=ws.title)

    if args.fill_metadata_search:
        print("Filling metadata from DuckDuckGo...")
        fill_metadata_from_duckduckgo(wb)

    if args.ddg_html:
        print("Filling metadata from DuckDuckGo HTML...")
        run_duckduckgo_html_metadata(ws)


    wb.save(workbook_path)
    print("Done.")


if __name__ == "__main__":
    main()
