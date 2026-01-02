#!/usr/bin/env python3

import os
import re
import time
import json
import sys
# import math
import unicodedata
import urllib.parse
import urllib.robotparser
import urllib.request
from collections import defaultdict, deque
from typing import List, Dict, Any, Tuple, Set

import tldextract
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

try:
    from openai import OpenAI
except Exception:
    OpenAI = None  # handled below

# Load Key.env first (if present), then fallback to .env.
# If python-dotenv isn't installed, do a simple fallback parser.
try:
    from dotenv import load_dotenv
    load_dotenv("Key.env")
    load_dotenv()
except Exception:
    # Lightweight fallback
    def _load_kv_file(path: str):
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#"):
                            continue
                        if "=" in line:
                            k, v = line.split("=", 1)
                            k = k.strip(); v = v.strip().strip('"').strip("'")
                            if k and v and k not in os.environ:
                                os.environ[k] = v
        except Exception:
            pass
    _load_kv_file("Key.env")
    _load_kv_file(".env")

# =====================
# CONFIG
# =====================
EXCEL_PATH = "/Users/patrickr/Documents/VS Code/Data enrichement.xlsx"  # input workbook (column A = websites)
INPUT_SHEET_NAME = "Analyse"  # explicit first sheet name
OUTPUT_SHEET_NAME = "Enriched Sheet"  # output sheet name (existing template)

HEADLESS = True
IGNORE_ROBOTS = True  # set True during testing to avoid robots.txt skips; set False to respect robots
NAV_TIMEOUT_MS = 25000 # max time per page navigation/load
PAGE_TIMEOUT_S = 60 # max time per page load/process
DOMAIN_TIMEOUT_S = 240 # max time per domain/site
STRICT_DOMAIN_ONLY = False  # TRUE = only keep emails that match the domain
MAX_PERSON_LINKS = 20  # max person-relevant links to follow per site
VERBOSE = True  # set True for detailed logs during testing
MAX_ROWS_TO_PROCESS = 0  # limit number of Analyse rows processed (None or 0 = unlimited)
POST_COMPANY_SLEEP_S = 3  # wait between companies to reduce API pressure during testing
LLM_MAX_PAGES = 2  # limit number of pages to send to LLM per site to reduce calls (0 = unlimited)

# Output language/tone (applies ONLY to Salutation (col H) and Icebreaker sentence (col L))
ENRICH_LANGUAGE = os.getenv("ENRICH_LANGUAGE", "es").strip().lower() # Allowed languages: de | fr | it | es
TONE = os.getenv("TONE", "formal").strip().lower() # Allowed tones: formal | informal


# LLM config
MODEL_NAME = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
MAX_LLM_CALLS_PER_MIN = 30  # simple throttle
MAX_PAGE_TEXT_CHARS = 12000  # truncate per-page text sent to LLM

TARGET_PERSON_KEYWORDS = [
    # English
    "contact", "imprint", "legal", "privacy", "about", "team", "careers", "people",
    # German
    "impressum", "anbieterkennzeichnung", "datenschutz", "ueber uns", "über uns", "kontakt", "team", "karriere", "wir",
    # Other EU languages often used
    "mentions legales", "mentions légales", "aviso legal", "note legali", "contatti", "empresa", "quienes somos",
]

EMAIL_RE = re.compile(r"([a-zA-Z0-9._%+\-]+)\s*@\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", re.IGNORECASE)
MAILTO_RE = re.compile(r"^mailto:(.+)", re.IGNORECASE)

# =====================
# HELPERS
# =====================

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c)) if s else s

# =====================
# Language / tone helpers (ONLY for Salutation + Icebreaker)
# =====================

ALLOWED_LANGS = {"de", "fr", "it", "es"}
ALLOWED_TONES = {"formal", "informal"}


def validate_language_and_tone(lang: str, tone: str) -> Tuple[str, str]:
    lang = (lang or "de").strip().lower()
    tone = (tone or "formal").strip().lower()
    if lang not in ALLOWED_LANGS:
        raise ValueError(f"Invalid ENRICH_LANGUAGE='{lang}'. Allowed: {sorted(ALLOWED_LANGS)}")
    if tone not in ALLOWED_TONES:
        raise ValueError(f"Invalid TONE='{tone}'. Allowed: {sorted(ALLOWED_TONES)}")
    return lang, tone


def make_salutation(first_name: str, last_name: str, gender: str, lang: str, tone: str) -> str:
    """Deterministic salutation generator.

    Applies ONLY to output column H.
    - formal: last-name based where possible; otherwise neutral formal.
    - informal: first-name based if available; otherwise neutral informal.
    """
    fn = normalize_spaces(first_name)
    ln = normalize_spaces(last_name)
    g = (gender or "unknown").strip().lower()
    lang, tone = validate_language_and_tone(lang, tone)

    # Informal
    if tone == "informal":
        if fn:
            return {
                "de": f"Hi {fn}",
                "fr": f"Salut {fn}",
                "it": f"Ciao {fn}",
                "es": f"Hola {fn}",
            }[lang]
        # neutral informal fallback (no name)
        return {
            "de": "Hallo",
            "fr": "Bonjour",
            "it": "Ciao",
            "es": "Hola",
        }[lang]

    # Formal
    if lang == "de":
        if g == "female" and ln:
            return f"Sehr geehrte Frau {ln}"
        if g == "male" and ln:
            return f"Sehr geehrter Herr {ln}"
        # neutral formal fallback
        return (f"Guten Tag {fn} {ln}".strip() if (fn or ln) else "Guten Tag")

    if lang == "fr":
        if g == "female" and ln:
            return f"Madame {ln}"
        if g == "male" and ln:
            return f"Monsieur {ln}"
        return (f"Bonjour {fn} {ln}".strip() if (fn or ln) else "Bonjour")

    if lang == "it":
        if g == "female" and ln:
            return f"Gentile Sig.ra {ln}"
        if g == "male" and ln:
            return f"Gentile Sig. {ln}"
        return (f"Buongiorno {fn} {ln}".strip() if (fn or ln) else "Buongiorno")

    # es
    if g == "female" and ln:
        return f"Estimada Sra. {ln}"
    if g == "male" and ln:
        return f"Estimado Sr. {ln}"
    return (f"Buenos días {fn} {ln}".strip() if (fn or ln) else "Buenos días")


def make_general_salutation(lang: str, tone: str) -> str:
    """Fallback greeting when no individual person is found.
    Applies ONLY to output column H.
    """
    lang, tone = validate_language_and_tone(lang, tone)
    if tone == "informal":
        return {
            "de": "Hallo",
            "fr": "Bonjour",
            "it": "Ciao",
            "es": "Hola",
        }[lang]
    return {
        "de": "Sehr geehrte Damen und Herren",
        "fr": "Madame, Monsieur",
        "it": "Gentili Signore e Signori",
        "es": "Estimados señores",
    }[lang]


def get_icebreaker_prompts(lang: str, tone: str) -> Tuple[str, str]:
    """Return (system_prompt, user_prompt_template) for the icebreaker (column L).

    Keeps other prompts unchanged.
    """
    lang, tone = validate_language_and_tone(lang, tone)

    # Language labels for the instruction line
    lang_label = {
        "de": "Deutsch",
        "fr": "Französisch",
        "it": "Italienisch",
        "es": "Spanisch",
    }[lang]

    # Tone instruction (kept short + explicit)
    tone_line = {
        "formal": {
            "de": "Ton: formell (Sie-Ansprache), professionell.",
            "fr": "Ton: formel (vouvoiement), professionnel.",
            "it": "Tono: formale (Lei), professionale.",
            "es": "Tono: formal (usted), profesional.",
        },
        "informal": {
            "de": "Ton: locker (Du-Ansprache), freundlich.",
            "fr": "Ton: informel (tutoiement), amical.",
            "it": "Tono: informale (tu), amichevole.",
            "es": "Tono: informal (tú), amistoso.",
        }
    }[tone][lang]

    # Main rules (translated minimally per language to reduce risk of mixed-language output)
    rules = {
        "de": (
            "Schreibe 1–2 sehr kurze, personalisierte und DSGVO-konforme Sätze auf Deutsch.\n"
            f"{tone_line}\n"
            "Regeln:\n"
            "- Schreibe wie eine persönliche Beobachtung eines Besuchers (z. B. 'Mir ist aufgefallen …'), nicht wie ein Slogan.\n"
            "- Beziehe dich auf ein konkretes Detail, vermeide Werbesprache und Superlative.\n"
            "- Verwende nur Inhalte aus dem bereitgestellten Startseitentext; keine Erfindungen.\n"
            "- Nenne NICHT den Unternehmensnamen oder Ort.\n"
            "- Maximal 150 Zeichen. Gib NUR den Satz bzw. die Sätze aus, kein JSON, keine Erklärungen."
        ),
        "fr": (
            "Écris 1–2 phrases très courtes, personnalisées et conformes au RGPD en français.\n"
            f"{tone_line}\n"
            "Règles :\n"
            "- Écris comme une observation personnelle d'un visiteur, pas comme un slogan.\n"
            "- Réfère-toi à un détail concret, évite le langage publicitaire et les superlatifs.\n"
            "- Utilise uniquement le texte fourni (page d'accueil) ; n'invente rien.\n"
            "- Ne cite PAS le nom de l'entreprise ni le lieu.\n"
            "- 150 caractères max. Donne UNIQUEMENT la/les phrase(s), pas de JSON, pas d'explications."
        ),
        "it": (
            "Scrivi 1–2 frasi molto brevi, personalizzate e conformi al GDPR in italiano.\n"
            f"{tone_line}\n"
            "Regole:\n"
            "- Scrivi come un'osservazione personale di un visitatore, non come uno slogan.\n"
            "- Cita un dettaglio concreto, evita linguaggio promozionale e superlativi.\n"
            "- Usa solo il testo fornito (homepage); non inventare.\n"
            "- Non menzionare il nome dell'azienda o la località.\n"
            "- Max 150 caratteri. Fornisci SOLO la/e frase/i, niente JSON, niente spiegazioni."
        ),
        "es": (
            "Escribe 1–2 frases muy cortas, personalizadas y conformes al RGPD en español.\n"
            f"{tone_line}\n"
            "Reglas:\n"
            "- Escribe como una observación personal de un visitante, no como un eslogan.\n"
            "- Menciona un detalle concreto; evita lenguaje publicitario y superlativos.\n"
            "- Usa solo el texto proporcionado (inicio); no inventes.\n"
            "- No menciones el nombre de la empresa ni la ubicación.\n"
            "- Máx. 150 caracteres. Devuelve SOLO la(s) frase(s), sin JSON ni explicaciones."
        ),
    }[lang]

    sys_prompt = rules
    user_tmpl = (
        "Domain: {site_url}\n"
        "Startseiten-Text (abgeschnitten):\n{homepage_text}\n\n"
        "Gib NUR den Satz/die Sätze aus. Kein JSON."
    )

    # Ensure language is explicitly stated even if user template is German
    sys_prompt = sys_prompt + f"\nSprache: {lang_label}."

    return sys_prompt, user_tmpl


def print_run_config():
    """Print key configs and a short safety confirmation banner."""
    lang, tone = validate_language_and_tone(ENRICH_LANGUAGE, TONE)
    cfg = {
        "EXCEL_PATH": EXCEL_PATH,
        "INPUT_SHEET_NAME": INPUT_SHEET_NAME,
        "OUTPUT_SHEET_NAME": OUTPUT_SHEET_NAME,
        "MODEL_NAME": MODEL_NAME,
        "ENRICH_LANGUAGE": lang,
        "TONE": tone,
        "HEADLESS": HEADLESS,
        "IGNORE_ROBOTS": IGNORE_ROBOTS,
        "STRICT_DOMAIN_ONLY": STRICT_DOMAIN_ONLY,
        "NAV_TIMEOUT_MS": NAV_TIMEOUT_MS,
        "PAGE_TIMEOUT_S": PAGE_TIMEOUT_S,
        "DOMAIN_TIMEOUT_S": DOMAIN_TIMEOUT_S,
        "MAX_ROWS_TO_PROCESS": MAX_ROWS_TO_PROCESS,
        "LLM_MAX_PAGES": LLM_MAX_PAGES,
        "MAX_PAGE_TEXT_CHARS": MAX_PAGE_TEXT_CHARS,
        "MAX_LLM_CALLS_PER_MIN": MAX_LLM_CALLS_PER_MIN,
        "POST_COMPANY_SLEEP_S": POST_COMPANY_SLEEP_S,
        "VERBOSE": VERBOSE,
    }
    print("\n[config] Run configuration")
    print("[config] " + "-" * 60)
    for k in [
        "EXCEL_PATH", "INPUT_SHEET_NAME", "OUTPUT_SHEET_NAME",
        "MODEL_NAME", "ENRICH_LANGUAGE", "TONE",
        "HEADLESS", "IGNORE_ROBOTS", "STRICT_DOMAIN_ONLY",
        "NAV_TIMEOUT_MS", "PAGE_TIMEOUT_S", "DOMAIN_TIMEOUT_S",
        "MAX_ROWS_TO_PROCESS", "LLM_MAX_PAGES",
        "MAX_PAGE_TEXT_CHARS", "MAX_LLM_CALLS_PER_MIN",
        "POST_COMPANY_SLEEP_S", "VERBOSE",
    ]:
        print(f"[config] {k}: {cfg[k]}")
    print("[config] " + "-" * 60 + "\n")


def countdown(seconds: int = 10):
    try:
        for remaining in range(int(seconds), 0, -1):
            print(f"[config] Starting in {remaining}s... (Ctrl+C to abort)")
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[abort] Cancelled by user.")
        sys.exit(0)



def keyword_hit(href: str, text: str) -> bool:
    hay = (href or "") + " " + (text or "")
    hay = strip_accents(hay).lower()
    return any(strip_accents(kw).lower() in hay for kw in TARGET_PERSON_KEYWORDS)

def _domain_core_from_url(url: str) -> str:
    try:
        ext = tldextract.extract(url)
        return re.sub(r'[^a-z0-9]', '', (ext.domain or '').lower())
    except Exception:
        return ''

def _domain_core_from_host(host: str) -> str:
    try:
        ext = tldextract.extract('http://' + (host or ''))
        return re.sub(r'[^a-z0-9]', '', (ext.domain or '').lower())
    except Exception:
        return ''


def same_company_domain(site_url: str, email_domain: str) -> bool:
    """Lenient matching to decide if an email domain likely belongs to the same business.

    Acceptance rules (to exclude only clearly unrelated partner companies):
    1) Same domain family different TLD: accept if second-level domain cores match
       (e.g., fischbach-miller.shop ↔ fischbach-miller.de).
    2) Parent brand or subdomain relationship: accept if one is a subdomain of the other
       or they share the same brand namespace (e.g., heads.volu-med.studio → volu-med.com).
    3) Official secondary domain (business ecosystem): accept if brand base matches even if
       structure differs (e.g., hairbyclaireshop.de → hairbyclaire.de).
    """
    try:
        import re as _re
        import tldextract as _tld

        def _norm_core(s: str) -> str:
            return _re.sub(r"[^a-z0-9]", "", (s or "").lower())

        def _extract(url_or_host: str):
            if not url_or_host:
                return "", "", "", ""
            ext = _tld.extract(url_or_host if _re.match(r"^https?://", str(url_or_host)) else ("http://" + str(url_or_host)))
            domain = (ext.domain or "").lower()
            suffix = (ext.suffix or "").lower()
            sub = (ext.subdomain or "").lower()
            host = ((domain + "." + suffix) if domain and suffix else domain)
            return sub, domain, suffix, host

        # Tokens considered generic add-ons rather than brand-identifiers
        GENERIC_TOKENS = {
            "shop","store","studio","wigs","hair","salon","beauty","clinic","praxis","co","group","gmbh","ug","mbh","ag","ev","kg","ohg","se","kgaa"
        }

        def _brand_tokens(domain: str):
            toks = [t for t in _re.split(r"[^a-z0-9]+", domain.lower()) if t]
            return [t for t in toks if t not in GENERIC_TOKENS]

        def _jaccard(a: set, b: set) -> float:
            if not a or not b:
                return 0.0
            inter = len(a & b)
            denom = max(len(a), len(b))
            return inter / denom if denom else 0.0

        s_sub, s_dom, s_suf, s_host = _extract(site_url)
        e_sub, e_dom, e_suf, e_host = _extract(email_domain)
        if not e_dom:
            return False

        # Rule 1: Same second-level domain (ignoring TLD and punctuation)
        if _norm_core(s_dom) and _norm_core(s_dom) == _norm_core(e_dom):
            return True

        # Rule 2a: Direct subdomain relationship on full host
        if s_host and e_host and (e_host.endswith(s_host) or s_host.endswith(e_host)):
            return True

        # Rule 2b/3: Brand-core similarity (namespace sharing / official secondary domain)
        s_tokens = _brand_tokens(s_dom)
        e_tokens = _brand_tokens(e_dom)
        if s_tokens and e_tokens:
            # If concatenated brand strings contain each other, accept
            s_join = "".join(s_tokens)
            e_join = "".join(e_tokens)
            if s_join in e_join or e_join in s_join:
                return True
            # If token Jaccard similarity is high, accept
            if _jaccard(set(s_tokens), set(e_tokens)) >= 0.6:
                return True

        return False
    except Exception:
        # Fallback to strict core equality if anything goes wrong
        site_core = _domain_core_from_url(site_url)
        email_core = _domain_core_from_host(email_domain)
        return bool(site_core and email_core and site_core == email_core)

def email_allowed(site_url: str, email_addr: str) -> bool:
    if not email_addr or '@' not in email_addr:
        return False
    edomain = email_addr.split('@')[-1].lower()
    if STRICT_DOMAIN_ONLY:
        return same_company_domain(site_url, edomain)
    return True

def html_preclean(html: str) -> str:
    if not html:
        return ""
    s = html
    s = re.sub(r"<style[^>]*>[\s\S]*?</style>", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"<script[^>]*>[\s\S]*?</script>", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"&#64;|&commat;|\(\s*at\s*\)|\[\s*at\s*\]|\{\s*at\s*\}", "@", s, flags=re.IGNORECASE)
    s = re.sub(r"\(\s*dot\s*\)|\[\s*dot\s*\]|\{\s*dot\s*\}", ".", s, flags=re.IGNORECASE)
    s = s.replace('>', ' ').replace('<', ' ').replace('&nbsp;', ' ').replace('&amp;', '&')
    return re.sub(r"\s+", " ", s).strip()


def clean_email_candidate(s: str):
    if not s:
        return None
    s = re.sub(r"[\(\[\{]\s*at\s*[\)\]\}]", "@", s, flags=re.IGNORECASE)
    s = re.sub(r"[\(\[\{]\s*dot\s*[\)\]\}]", ".", s, flags=re.IGNORECASE)
    s = re.sub(r"\bat\b", "@", s, flags=re.IGNORECASE)
    s = re.sub(r"\bdot\b", ".", s, flags=re.IGNORECASE)
    s = s.replace(">", " ").replace("<", " ").replace("&nbsp;", " ").replace("&amp;", "&")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.split(r"[ \t\n\r;,\?]", s)[0]
    m = EMAIL_RE.search(s)
    return m.group(0).lower() if m else None


def extract_emails_from_text(text: str) -> Set[str]:
    emails = set()
    for m in EMAIL_RE.finditer(text):
        emails.add((m.group(1) + '@' + m.group(2)).lower())
    for chunk in re.findall(r"[\w.\-\[\]\(\)\{@\s]+(?:at|@|&#64;|&commat;)[\w.\-]+\.\w{2,}", text, flags=re.IGNORECASE):
        e = clean_email_candidate(chunk)
        if e:
            emails.add(e)
    return emails


def load_robots(base_url):
    rp = urllib.robotparser.RobotFileParser()
    robots_url = urllib.parse.urljoin(
        f"{urllib.parse.urlparse(base_url).scheme}://{urllib.parse.urlparse(base_url).netloc}",
        "/robots.txt",
    )
    try:
        with urllib.request.urlopen(robots_url, timeout=7) as resp:
            content = resp.read().decode("utf-8", errors="ignore")
        rp.parse(content.splitlines())
        print(f"[robots] Loaded: {robots_url}")
        return rp
    except Exception as e:
        print(f"[robots] Skip robots (error: {type(e).__name__}) -> {robots_url}")
        return None


def allowed_by_robots(rp, url):
    return True if IGNORE_ROBOTS or not rp else rp.can_fetch("PeopleEnricher/Playwright", url)


# =====================
# Playwright helpers
# =====================

def get_rendered_html(page, url) -> Tuple[str, str]:
    page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    return page.content(), page.url


def discover_person_links(page, base_url) -> List[str]:
    elements = page.eval_on_selector_all(
        "a[href]",
        "els => els.map(e => ({ href: e.getAttribute('href'), text: e.innerText }))"
    )
    abs_links: List[str] = []
    for el in elements:
        href = (el.get('href') or '').strip()
        text = normalize_spaces(el.get('text'))
        if not href:
            continue
        low = href.lower()
        if low.startswith('javascript:'):
            continue
        if low.startswith('mailto:'):
            continue
        try:
            full = urllib.parse.urljoin(base_url, href.split('#', 1)[0])
        except Exception:
            continue
        if keyword_hit(full, text):
            abs_links.append(full)
    # de-dup preserve order
    seen = set()
    kept = []
    for l in abs_links:
        if l not in seen:
            kept.append(l)
            seen.add(l)
    return kept[:MAX_PERSON_LINKS]


def prioritize_person_links(links: List[str]) -> List[str]:
    # Simple keyword-based scoring; higher is better
    weights = [
        (r'impressum|imprint|anbieterkennzeichnung', 100),
        (r'kontakt|contact', 80),
        (r'team|ueber-uns|über-uns|about', 70),
        (r'privacy|datenschutz', 30),
        (r'legal', 25),
    ]
    def score(url: str) -> int:
        s = url.lower()
        total = 0
        for pat, w in weights:
            if re.search(pat, s):
                total += w
        return total
    return sorted(links, key=score, reverse=True)


def extract_mailto_emails_from_dom(page, site_url: str) -> Set[str]:
    items = page.eval_on_selector_all('a[href^="mailto:"]', "els => els.map(e => e.getAttribute('href'))")
    out = set()
    for href in items or []:
        try:
            cand = (href or '').split(':', 1)[1].split('?', 1)[0]
            e = clean_email_candidate(cand)
            if e and email_allowed(site_url, e):
                out.add(e.lower())
        except Exception:
            continue
    return out

# =====================
# LLM client with simple rate limiter and retries
# =====================
class LLMRateLimiter:
    def __init__(self, max_calls_per_min: int):
        self.max_calls = max(1, int(max_calls_per_min))
        self.calls = deque()

    def wait_for_slot(self):
        now = time.time()
        while self.calls and now - self.calls[0] > 60.0:
            self.calls.popleft()
        if len(self.calls) >= self.max_calls:
            sleep_s = 60.0 - (now - self.calls[0]) + 0.01
            sleep_s = max(0.5, min(30.0, sleep_s))
            if VERBOSE:
                print(f"[llm] Throttle: sleeping {sleep_s:.1f}s")
            time.sleep(sleep_s)
        self.calls.append(time.time())


class LLMClient:
    def __init__(self, model: str, rate_limiter: LLMRateLimiter):
        if OpenAI is None:
            raise RuntimeError("OpenAI SDK not installed. Run: pip install openai")
        if not os.getenv("OPENAI_API_KEY"):
            raise RuntimeError("OPENAI_API_KEY not set in environment.")
        self.client = OpenAI()
        self.model = model
        self.rate = rate_limiter

    @retry(stop=stop_after_attempt(4), wait=wait_exponential(multiplier=1, min=1, max=12),
           retry=retry_if_exception_type(Exception))
    def chat_json(self, system_prompt: str, user_prompt: str) -> Dict[str, Any]:
        self.rate.wait_for_slot()
        resp = self.client.chat.completions.create(
            model=self.model,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.1,
        )
        content = resp.choices[0].message.content or "{}"
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            # try to salvage JSON
            cleaned = content.strip()
            start = cleaned.find('{')
            end = cleaned.rfind('}')
            if start != -1 and end != -1 and end > start:
                return json.loads(cleaned[start:end+1])
            raise


@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=1, max=8),
       retry=retry_if_exception_type(Exception))
def llm_chat_text(llm: LLMClient, system_prompt: str, user_prompt: str) -> str:
    # Simple text response without JSON enforcement; used as a robust fallback
    llm.rate.wait_for_slot()
    resp = llm.client.chat.completions.create(
        model=llm.model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.2,
    )
    return (resp.choices[0].message.content or "").strip()

# =====================
# LLM prompts
# =====================
PERSON_SYS = (
    "Du bist ein Extraktionsassistent für deutschsprachige Webseiten (Impressum/Kontakt/Team). Arbeite ausschließlich mit den bereitgestellten Kontextfenstern.\n"
    "Regeln:\n"
    "- Extrahiere NUR Entscheidungsträger/innen mit klaren Rollen: Verantwortlich/inhaltlich verantwortlich, Inhaber/Inhaberin, Geschäftsführer/Geschäftsführerin, Geschäftsführung, Kontoinhaber, Leitung.\n"
    "- Wenn keine Person mit einer dieser Rollen eindeutig genannt ist, gib 'people': [] zurück (nicht raten, nicht erzwingen).\n"
    "- Erkenne Personennamen (2–4 Tokens, Großschreibung; Umlaut/ß/Bindestriche erlaubt). Bewahre Diakritika.\n"
    "- Behandle Rechtsformzusätze wie 'GmbH', 'UG', 'mbH', 'AG', 'e.V.', 'GbR', 'KG', 'OHG', 'SE' niemals als Vor- oder Nachnamen.\n"
    "- Erwähne oder extrahiere KEINE Firmen- oder Dienstleister-Namen. Nur natürliche Personen.\n"
    "- Gib nur Informationen zurück, die in den Fenstern vorkommen. Keine Halluzinationen.\n"
    "- E-Mail nur, wenn sie in 'Bekannte E-Mails' oder im Fenstertext vorkommt.\n"
    "- Felder pro Person: first_name, last_name, role, gender, email, salutation.\n"
    "- gender: male | female | unknown (nur wenn klar erkennbar).\n"
    "- salutation-Regeln: female+Nachname => 'Sehr geehrte Frau [Nachname]'; male+Nachname => 'Sehr geehrter Herr [Nachname]'; female ohne Nachname => 'Liebe [Vorname]'; male ohne Nachname => 'Lieber [Vorname]'; sonst ''.\n"
    "Format: JSON-Objekt mit Feld 'people' = Liste von Personenobjekten.\n"
)
PERSON_USER_TMPL = (
    "Website: {site_url}\n"
    "Seite: {page_url}\n"
    "Bekannte E-Mails: {emails}\n\n"
    "Kontextfenster (Leerzeichen normalisiert, Diakritika erhalten):\n{windows}\n\n"
    "Gib ein JSON-Objekt mit: people: [{{first_name, last_name, role, gender, email, salutation}}].\n"
    "- Bevorzuge die am höchsten priorisierten Labels. Wenn mehrere Personen vorkommen, nenne bis zu 3 mit klarer Rolle.\n"
)

COMPANY_SYS = (
    "Du analysierst die Startseite eines Unternehmens und gibst eine knappe, sachliche Zusammenfassung (max. 25 Wörter).\n"
    "Gib ein JSON-Objekt mit GENAU einem Feld: company_summary.\n"
    "Nur aus Text ableiten, keine Vermutungen. Sprache: Deutsch.\n"
)

COMPANY_USER_TMPL = (
    "Domain: {site_url}\n"
    "Startseiten-Text (abgeschnitten):\n{homepage_text}\n\n"
    "Gib ein JSON-Objekt mit dem Feld company_summary."
)


ICEBREAKER_SYS_PLAIN = (
    "Schreibe 1–2 sehr kurze, personalisierte und DSGVO-konforme Sätze auf Deutsch.\n"
    "Regeln:\n"
    "- Schreibe wie eine persönliche Beobachtung eines Besuchers (z. B. 'Mir ist aufgefallen …', 'Ich finde super …'), nicht wie ein Slogan.\n"
    "- Beziehe dich auf ein konkretes Detail (Leistungen, Spezialisierung, Besonderheiten), vermeide Werbesprache und Superlative.\n"
    "- Verwende nur Inhalte aus dem bereitgestellten Startseitentext; keine Erfindungen.\n"
    "- Nenne NICHT den Unternehmensnamen oder Ort.\n"
    "- Maximal 150 Zeichen. Gib NUR den Satz bzw. die Sätze aus, kein JSON, keine Erklärungen."
)

ICEBREAKER_USER_PLAIN = (
    "Domain: {site_url}\n"
    "Startseiten-Text (abgeschnitten):\n{homepage_text}\n\n"
    "Gib NUR den Satz/die Sätze aus. Kein JSON."
)


# =====================
# Kontextfenster & Fallback-Personenerkennung
# =====================

ANCHOR_TERMS = [
    r"\bverantwortlich(?:e|er)?\b",
    r"\binhaber(?:in)?\b",
    r"\bgeschäftsführer(?:in)?\b",
    r"\bansprechpartner(?:in)?\b",
    r"\bkontoinhaber\b",
    r"\bgeschäftsführung\b",
    r"\bleitung\b",
    r"\b(unser|ihr|das)\s+team\b",
    r"\bteam\b",
    r"\bimpressum\b",
    r"\bdatenschutz\b",
    r"\bkontakt\b",
    r"\biban\b",
]

def _normalize_nfc_spaces(text: str) -> str:
    t = unicodedata.normalize("NFC", text or "")
    t = re.sub(r"\s+", " ", t)
    return t.strip()

def _find_spans(rx: re.Pattern, text: str):
    return [m.span() for m in rx.finditer(text)]

def _merge_ranges(ranges: List[Tuple[int, int]], radius: int, length: int) -> List[Tuple[int, int]]:
    expanded = []
    for s, e in ranges:
        s2 = max(0, s - radius)
        e2 = min(length, e + radius)
        expanded.append((s2, e2))
    expanded.sort()
    merged: List[Tuple[int, int]] = []
    for s, e in expanded:
        if not merged or s > merged[-1][1] + 10:
            merged.append([s, e])
        else:
            merged[-1][1] = max(merged[-1][1], e)
    return [(s, e) for s, e in merged]

def build_context_windows(text: str, radius: int = 600, max_windows: int = 20) -> List[str]:
    t = _normalize_nfc_spaces(text)
    spans: List[Tuple[int, int]] = []
    spans.extend([m.span() for m in EMAIL_RE.finditer(t)])
    for pat in ANCHOR_TERMS:
        spans.extend([m.span() for m in re.finditer(pat, t, flags=re.IGNORECASE)])
    if not spans:
        return [t[:2000]] if t else []
    ranges = _merge_ranges(spans, radius, len(t))
    # Score ranges: prioritize those containing high-signal labels and emails
    def score_range(r: Tuple[int, int]) -> Tuple[int, int]:
        s, e = r
        w = t[s:e]
        pri = 0
        if re.search(r"\bverantwortlich", w, re.I): pri += 5
        if re.search(r"\binhaber", w, re.I): pri += 4
        if re.search(r"\bgeschäftsführer", w, re.I): pri += 4
        if re.search(r"\bkontoinhaber", w, re.I): pri += 3
        if re.search(EMAIL_RE, w): pri += 1
        return (-pri, e - s)
    ranges.sort(key=score_range)
    selected = ranges[:max_windows]
    return [t[s:e] for s, e in selected]

NAME_RE = r"([A-ZÄÖÜ][\wÄÖÜäöüß'.-]+(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß'.-]+){1,3})"
RESP_PATTERNS = [
    (re.compile(r"(?i)\bverantwortlich(?:e|er)?\b[:\- ]+" + NAME_RE), "Verantwortlich"),
    (re.compile(r"(?i)\binhaber(?:in)?\b[:\- ]+" + NAME_RE), "Inhaber/in"),
    (re.compile(r"(?i)\bgeschäftsführer(?:in)?\b[:\- ]+" + NAME_RE), "Geschäftsführer/in"),
    (re.compile(r"(?i)\bgeschäftsführung\b[:\- ]+" + NAME_RE), "Geschäftsführung"),
    (re.compile(r"(?i)\bkontoinhaber\b[:\- ]+" + NAME_RE), "Kontoinhaber"),
]

def fallback_person(windows: List[str]) -> Tuple[str | None, str | None, str | None]:
    for w in windows or []:
        w2 = _normalize_nfc_spaces(w)
        for rx, label in RESP_PATTERNS:
            m = rx.search(w2)
            if m:
                name = m.group(1).strip(" .,-")
                # Guard against companies/service providers or non-name tokens
                fn_tmp, ln_tmp = split_name_simple(name)
                if not looks_like_person_name(fn_tmp, ln_tmp):
                    continue
                i, j = m.span(1)
                evid = w2[max(0, i-40):min(len(w2), j+40)]
                return name, label, evid
    return None, None, None

def split_name_simple(full_name: str) -> Tuple[str, str]:
    parts = [p for p in (full_name or "").strip().split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]

LEGAL_FORMS = {"gmbh", "mbh", "ug", "ag", "e.v.", "ev", "gbr", "kg", "ohg", "se", "kgaa", "ek", "e.k.", "sarl", "s.a.r.l.", "srl", "s.p.a.", "spa"}
PROVIDER_BLACKLIST = {"sendinblue", "brevo", "mailchimp", "mailerlite", "hubspot", "google", "meta", "facebook", "instagram", "linkedin", "twitter", "x", "stripe", "paypal", "klarna", "mailjet", "activecampaign", "constant", "brevo"}

# Decision-maker role matcher (German + English variants)
DECISION_ROLE_RE = re.compile(r"(?i)verantwortlich|inhaber|inhaberin|geschäftsführer|geschaeftsfuehrer|geschäftsführerin|geschaeftsfuehrerin|geschäftsführung|geschaeftsfuehrung|kontoinhaber|leitung|owner|founder|ceo")

# German stopwords and pronouns that must not be treated as names
STOPWORD_TOKENS = {
    "für","der","die","das","und","oder","im","in","am","an","beim","von","vom","zu","zur","zum","mit","ohne","über","ueber","unter","zwischen","gegen","nach","vor","seit","bis","ab","auch",
    "ihr","ihre","ihrer","ihrem","ihren","euer","eure","eurer","eurem","euren","unser","unsere","unserer","unserem","unseren",
}

def _starts_upper(token: str) -> bool:
    return bool(re.match(r"^[A-ZÄÖÜ][\wÄÖÜäöüß'.-]+$", token or ""))

def looks_like_person_name(fn: str, ln: str) -> bool:
    if not fn or not ln:
        return False
    # Filter legal forms / providers
    if is_legal_name(fn, ln) or is_blacklisted_name(fn, ln):
        return False
    # Token checks
    fn_l = fn.strip().strip('.').lower()
    ln_l = ln.strip().strip('.').lower()
    if fn_l in STOPWORD_TOKENS or ln_l in STOPWORD_TOKENS:
        return False
    if any(ch.isdigit() for ch in fn) or any(ch.isdigit() for ch in ln):
        return False
    # Require typical capitalization (allow particles like 'von', 'de', 'van' as first token)
    allowed_particles = {"von","van","de","da","di","zu","zur","zum"}
    fn_ok = _starts_upper(fn) or fn_l in allowed_particles
    ln_ok = _starts_upper(ln)
    return fn_ok and ln_ok

def _tokenize_name_tokens(*names: str) -> List[str]:
    toks: List[str] = []
    for n in names:
        for t in (n or "").replace(",", " ").split():
            t2 = t.strip().strip(".").lower()
            if t2:
                toks.append(t2)
    return toks

def is_legal_name(fn: str, ln: str) -> bool:
    toks = _tokenize_name_tokens(fn, ln)
    return any(t in LEGAL_FORMS for t in toks)

def is_blacklisted_name(fn: str, ln: str) -> bool:
    toks = _tokenize_name_tokens(fn, ln)
    return any(t in PROVIDER_BLACKLIST for t in toks)

# =====================
# Excel I/O
# =====================

def open_workbook():
    from openpyxl import load_workbook
    import os
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")
    return load_workbook(EXCEL_PATH)


def read_input_urls(wb) -> List[Tuple[str, str]]:
    ws = wb[wb.sheetnames[0]] if INPUT_SHEET_NAME is None else wb[INPUT_SHEET_NAME]
    sites: List[Tuple[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        company = (row[0] or "").strip() if row and len(row) > 0 and row[0] else ""
        website = (row[1] or "").strip() if row and len(row) > 1 and row[1] else ""
        if not website:
            continue
        website = website if website.startswith("http") else "https://" + website
        sites.append((company, website))
    return sites


def ensure_people_sheet(wb):
    # Use existing template only; do not create new sheets or headers
    if OUTPUT_SHEET_NAME in wb.sheetnames:
        return wb[OUTPUT_SHEET_NAME]
    raise ValueError(f"Output sheet not found: {OUTPUT_SHEET_NAME}")


def append_people_rows(ws, rows: List[List[Any]]):
    for r in rows:
        ws.append(r)

# Header-aware append so we always write to the next truly free row (ignoring Excel's used-range artifacts)

def get_header_map(ws) -> Dict[str, int]:
    headers = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        key = str(cell.value or "").strip()
        if key:
            headers[key] = col_idx
    return headers


def _row_has_any_value(ws, row_idx: int, cols: List[int]) -> bool:
    for c in cols:
        v = ws.cell(row=row_idx, column=c).value
        if v is not None and str(v).strip() != "":
            return True
    return False


def find_next_value_row(ws, candidate_cols: List[int]) -> int:
    last = 1
    maxr = ws.max_row or 1
    for r in range(maxr, 1, -1):
        if _row_has_any_value(ws, r, candidate_cols):
            last = r
            break
    return last + 1


def append_rows_by_headers(ws, header_map: Dict[str, int], row_dicts: List[Dict[str, Any]]):
    # Determine columns that define a used row (prefer Company Name + Website URL)
    key_cols = []
    for name in ("Company Name", "Website URL"):
        if name in header_map:
            key_cols.append(header_map[name])
    if not key_cols:
        key_cols = list(header_map.values())

    for row_data in row_dicts:
        target_row = find_next_value_row(ws, key_cols)
        for k, v in row_data.items():
            if k in header_map:
                ws.cell(row=target_row, column=header_map[k], value=v)


def excel_file_locked(path: str) -> bool:
    """Return True only if we have strong evidence the workbook is open elsewhere.

    Rules:
    - If advisory lock (flock) is available and succeeds, treat as NOT locked, even if a ~$
      lock file exists (likely stale). Attempt to remove the lock file.
    - If advisory lock fails with BlockingIOError, treat as locked.
    - If flock is not available, fall back to lock-file presence, but consider files older
      than 5 minutes as stale and remove them.
    - Environment override: set EXCEL_IGNORE_LOCKS=1 to skip this check entirely.
    """
    if os.getenv("EXCEL_IGNORE_LOCKS") == "1":
        if VERBOSE:
            print("[excel] EXCEL_IGNORE_LOCKS=1 — skipping lock checks.")
        return False

    try:
        dirn, base = os.path.split(path)
        lock_path = os.path.join(dirn, "~$" + base)
        has_lock_file = os.path.exists(lock_path)

        # Try advisory lock first (best signal on POSIX)
        try:
            import fcntl  # only on POSIX
            with open(path, "rb") as f:
                try:
                    fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    # Lock acquired => not locked by another process.
                    fcntl.flock(f.fileno(), fcntl.LOCK_UN)
                    if has_lock_file:
                        # If we can lock, the lock file is stale — try to remove it.
                        try:
                            os.remove(lock_path)
                            if VERBOSE:
                                print(f"[excel] Removed stale Excel lock file: {lock_path}")
                        except Exception:
                            pass
                    return False
                except BlockingIOError:
                    if VERBOSE:
                        print("[excel] Advisory lock in use (BlockingIOError).")
                    return True
        except Exception:
            # No flock available or failed for non-lock reasons — fall back to lock file heuristic
            pass

        # Fallback when flock isn't usable: rely on lock file recency
        if has_lock_file:
            try:
                age = time.time() - os.path.getmtime(lock_path)
            except Exception:
                age = 0
            if age > 300:  # >5 minutes => stale; try remove and continue
                try:
                    os.remove(lock_path)
                    if VERBOSE:
                        print(f"[excel] Removed stale Excel lock file (>5min): {lock_path}")
                except Exception:
                    pass
                return False
            if VERBOSE:
                print(f"[excel] Detected Excel lock file: {lock_path}")
            return True

        return False

    except Exception:
        # Any unexpected issue -> assume not locked, let later save path report real errors.
        return False


def norm_url(u: str) -> str:
    u = (u or "").strip().lower()
    try:
        p = urllib.parse.urlparse(u if u.startswith("http") else "https://" + u)
        host = p.netloc.lstrip("www.")
        return host
    except Exception:
        return u


def read_existing_websites(wb) -> Set[str]:
    existing: Set[str] = set()
    if OUTPUT_SHEET_NAME in wb.sheetnames:
        ws = wb[OUTPUT_SHEET_NAME]
        header_map = get_header_map(ws)
        w_col = header_map.get("Website URL")
        if w_col:
            for row_idx in range(2, (ws.max_row or 1) + 1):
                website = str(ws.cell(row=row_idx, column=w_col).value or "").strip()
                if website:
                    existing.add(norm_url(website))
    return existing


# =====================
# Core pipeline per site
# =====================

def extract_page_text_and_emails(html: str) -> Tuple[str, Set[str]]:
    cleaned = html_preclean(html)
    soup = BeautifulSoup(cleaned, 'html.parser')
    text = soup.get_text(" ")
    text = normalize_spaces(text)
    emails = extract_emails_from_text(text)
    return text, emails


def infer_company_name(site_url: str) -> str:
    es = tldextract.extract(site_url)
    return es.domain.capitalize() if es and es.domain else site_url


def process_site(site_url: str, company_name: str, llm: LLMClient, wb) -> int:
    # returns number of people written
    from openpyxl.utils import get_column_letter

    if VERBOSE:
        print(f"[process] Company='{(company_name or '').strip()}', URL='{site_url}'")

    rp = load_robots(site_url)



    # Ensure we can write at least a company-only row even if robots disallow
    people_ws = ensure_people_sheet(wb)
    header_map = get_header_map(people_ws)  # add this line
    comp_name = (company_name or infer_company_name(site_url)).strip()

    merged_emails: Set[str] = set() 

    # If robots disallow and we're respecting robots, still write a company-only row
    if not allowed_by_robots(rp, site_url):
        print(f"[skip] robots.txt disallows: {site_url}")
        empty_dict = {
            "Company Name": comp_name,
            "Website URL": site_url,
            "Company summary": "",
            "Icebreaker sentence": "",
        }
        append_rows_by_headers(people_ws, header_map, [empty_dict])
        wb.save(EXCEL_PATH)
        print(f"[excel] Wrote company-only row due to robots.txt for {site_url}")
        return 0

    t0 = time.monotonic()
    people_rows: List[Dict[str, Any]] = []
    homepage_text = ""
    found_total = 0
    company_summary = ""
    seen_people_keys: Set[str] = set()  # dedupe across pages

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(user_agent="PeopleEnricher/Playwright", ignore_https_errors=True)
        def _router(route):
            req = route.request
            if req.resource_type in ("image", "media", "font"):
                return route.abort()
            bad = ("doubleclick.net", "googletagmanager.com", "googlesyndication.com", "facebook.net", "twitter.com", "hotjar", "cloudflareinsights")
            if any(b in req.url for b in bad):
                return route.abort()
            return route.continue_()
        context.route("**/*", _router)
        context.set_default_navigation_timeout(NAV_TIMEOUT_MS)
        context.set_default_timeout(NAV_TIMEOUT_MS)
        page = context.new_page()

        # Load homepage
        try:
            html0, final0 = get_rendered_html(page, site_url)
            homepage_text, homepage_emails = extract_page_text_and_emails(html0)
            # filter and add
            homepage_emails = {e for e in homepage_emails if email_allowed(site_url, e)}
            mailtos_home = extract_mailto_emails_from_dom(page, site_url)
            merged_emails.update(homepage_emails)
            merged_emails.update(mailtos_home)
            person_links = discover_person_links(page, final0)
            person_links = prioritize_person_links(person_links)
            if VERBOSE:
                print(f"[discover] {site_url} -> {len(person_links)} person pages; homepage emails: {len(homepage_emails|mailtos_home)}")
            # Phase 1: pre-scan target pages for emails only (no LLM)
            target_pages = [site_url] + person_links if person_links else [site_url]
            if VERBOSE and target_pages:
                print(f"[emails] {site_url}: scanning {len(target_pages)} page(s) for emails")
            for url in target_pages:
                if not allowed_by_robots(rp, url):
                    if VERBOSE:
                        print(f"[skip] robots.txt (subpage): {url}")
                    continue
                try:
                    html_tmp, final_tmp = get_rendered_html(page, url)
                    page_text_tmp, page_emails_tmp = extract_page_text_and_emails(html_tmp)
                    page_emails_tmp = {e for e in page_emails_tmp if email_allowed(site_url, e)}
                    mailtos_tmp = extract_mailto_emails_from_dom(page, site_url)
                    merged_emails.update(page_emails_tmp)
                    merged_emails.update(mailtos_tmp)
                except Exception:
                    continue
            if not merged_emails:
                empty_dict = {"Company Name": comp_name, "Website URL": site_url, "Company summary": "", "Icebreaker sentence": ""}
                append_rows_by_headers(people_ws, header_map, [empty_dict])
                try:
                    wb.save(EXCEL_PATH)
                except PermissionError:
                    print(f"[excel] Save failed: file appears open/locked: {EXCEL_PATH}")
                    sys.exit(1)
                print(f"[excel] No emails found. Wrote company-only row for {site_url}")
                return 0
        except PWTimeout:
            print(f"[timeout] Initial load timed out: {site_url}")
            browser.close()
            empty_dict = {"Company Name": comp_name, "Website URL": site_url, "Company summary": "", "Icebreaker sentence": ""}
            append_rows_by_headers(people_ws, header_map, [empty_dict])
            try:
                wb.save(EXCEL_PATH)
            except PermissionError:
                print(f"[excel] Save failed: file appears open/locked: {EXCEL_PATH}")
                sys.exit(1)
            print(f"[excel] Wrote company-only row after timeout for {site_url}")
            return 0
        except Exception as e:
            print(f"[error] Initial fetch: {e}")
            browser.close()
            empty_dict = {"Company Name": comp_name, "Website URL": site_url, "Company summary": "", "Icebreaker sentence": ""}
            append_rows_by_headers(people_ws, header_map, [empty_dict])
            try:
                wb.save(EXCEL_PATH)
            except PermissionError:
                print(f"[excel] Save failed: file appears open/locked: {EXCEL_PATH}")
                sys.exit(1)
            print(f"[excel] Wrote company-only row after error for {site_url}")
            return 0

        # Prepare homepage snippet once
        h_text = homepage_text[:MAX_PAGE_TEXT_CHARS]

        # Company analysis (summary only)
        try:
            company_payload = COMPANY_USER_TMPL.format(site_url=site_url, homepage_text=h_text)
            company_json = llm.chat_json(COMPANY_SYS, company_payload)
            if isinstance(company_json, dict):
                company_summary = normalize_spaces(str(company_json.get("company_summary", "")))
            if VERBOSE:
                print(f"[company] Summary: {company_summary[:120]}")
        except Exception as e:
            if VERBOSE:
                print(f"[warn] Company analysis LLM failed: {e}")

        # Icebreaker sentence (plain text, no JSON)
        icebreaker_sentence = ""
        try:
            ice_sys, ice_user_tmpl = get_icebreaker_prompts(ENRICH_LANGUAGE, TONE)
            raw = llm_chat_text(
                llm,
                ice_sys,
                ice_user_tmpl.format(site_url=site_url, homepage_text=h_text),
            )
            icebreaker_sentence = normalize_spaces(raw.strip().strip('"').strip("'"))
        except Exception:
            if VERBOSE:
                print("[info] Icebreaker text fetch failed; leaving empty.")
            icebreaker_sentence = ""

        # If no person links found, also analyze homepage as fallback
        target_pages = [site_url] + person_links if person_links else [site_url]
        if VERBOSE and target_pages:
            print(f"[crawl] {site_url}: processing {len(target_pages)} page(s)")

        for idx, url in enumerate(target_pages, 1):
            if time.monotonic() - t0 > DOMAIN_TIMEOUT_S:
                print(f"[watchdog] Domain timeout reached for {site_url}")
                break
            if not allowed_by_robots(rp, url):
                if VERBOSE:
                    print(f"[skip] robots.txt (subpage): {url}")
                continue
            try:
                html, final = get_rendered_html(page, url)
                page_text, page_emails = extract_page_text_and_emails(html)
                if VERBOSE:
                    print(f"[page] {idx}/{len(target_pages)} {final} — emails found: {len(page_emails)}")
                # Optionally limit LLM calls to first N pages per site
                do_llm = (LLM_MAX_PAGES <= 0) or (idx <= LLM_MAX_PAGES)
                people = []
                if do_llm:
                    windows = build_context_windows(page_text, radius=600, max_windows=15)
                    if VERBOSE:
                        print(f"[windows] {final}: built {len(windows)} window(s)")
                    windows_str = "\n".join(f"{i+1}) {w}" for i, w in enumerate(windows)) if windows else page_text[:MAX_PAGE_TEXT_CHARS]
                    payload = PERSON_USER_TMPL.format(
                        site_url=site_url,
                        page_url=final,
                        emails=", ".join(sorted(page_emails)) if page_emails else "(keine)",
                        windows=windows_str
                    )
                    out = llm.chat_json(PERSON_SYS, payload)
                    people = out.get("people", []) if isinstance(out, dict) else []
                    if not isinstance(people, list):
                        people = []
                    if not people:
                        fp, flabel, fev = fallback_person(windows)
                        if fp:
                            fn_f, ln_f = split_name_simple(fp)
                            people = [{
                                "first_name": fn_f,
                                "last_name": ln_f,
                                "role": flabel or "",
                                "gender": "unknown",
                                "email": "",
                                "salutation": "",
                            }]
                            if VERBOSE:
                                print(f"[fallback] {final}: matched responsible person via regex: {fp} ({flabel})")
                    if VERBOSE:
                        print(f"[llm] {final}: extracted people = {len(people)}")
                for p_obj in people:
                    try:
                        fn = normalize_spaces(str(p_obj.get("first_name", "")))
                        ln = normalize_spaces(str(p_obj.get("last_name", "")))
                        if not fn or not ln:
                            continue
                        role = normalize_spaces(str(p_obj.get("role") or p_obj.get("title") or p_obj.get("position") or ""))
                        email = normalize_spaces(str(p_obj.get("email", ""))).lower()
                        # Accept any syntactically valid email; only restrict by domain if STRICT_DOMAIN_ONLY is enabled.
                        if email:
                            if not EMAIL_RE.search(email):
                                email = ""
                            elif STRICT_DOMAIN_ONLY:
                                edomain = email.split('@')[-1].lower()
                                if not same_company_domain(site_url, edomain):
                                    email = ""
                        gender = str(p_obj.get("gender", "unknown")).lower()
                        sal = ""
                        # Strong validation: name must look like a person and role must indicate decision-making
                        if not looks_like_person_name(fn, ln):
                            if VERBOSE:
                                print(f"[filter] Rejecting unlikely name: '{fn} {ln}' role='{role}'")
                            continue
                        if not DECISION_ROLE_RE.search(role or ""):
                            if VERBOSE:
                                print(f"[filter] Rejecting non-decision role: '{fn} {ln}' role='{role}'")
                            continue
                        # Deterministic salutation (language + tone)
                        sal = make_salutation(fn, ln, gender, ENRICH_LANGUAGE, TONE)
                        key = (fn.lower(), ln.lower(), email)
                        if key in seen_people_keys:
                            continue
                        seen_people_keys.add(key)
                        people_rows.append({
                            "Company Name": comp_name,
                            "Website URL": site_url,
                            "Company summary": company_summary,
                            "Job title": role,
                            "First name": fn,
                            "Last name": ln,
                            "Gender": gender,
                            "Salutation": sal,
                            "Emails": email,
                            "Icebreaker sentence": icebreaker_sentence,
                        })
                        found_total += 1
                    except Exception:
                        continue
            except PWTimeout:
                print(f"[timeout] Page load timed out: {url}")
            except Exception as e:
                print(f"[error] Fetch {url}: {e}")

        browser.close()

    # Append to Excel with before/after row counts (true last-used rows)
    try:
        # compute true last-used row using key columns
        key_cols = []
        for name in ("Company Name", "Website URL"):
            if name in header_map:
                key_cols.append(header_map[name])
        if not key_cols:
            key_cols = list(header_map.values())
        before_rows = find_next_value_row(people_ws, key_cols) - 1

        if people_rows:
            append_rows_by_headers(people_ws, header_map, people_rows)
            try:
                wb.save(EXCEL_PATH)
            except PermissionError:
                print(f"[excel] Save failed: file appears open/locked: {EXCEL_PATH}")
                sys.exit(1)
            after_rows = find_next_value_row(people_ws, key_cols) - 1
            print(f"[excel] {EXCEL_PATH} :: '{OUTPUT_SHEET_NAME}' — wrote {len(people_rows)} row(s), people={len(people_rows)}; rows before={before_rows}, after={after_rows}")
        else:
            # Even if nothing else found, write company + website row with all emails and general salutation
            empty_dict = {
                "Company Name": comp_name,
                "Website URL": site_url,
                "Company summary": company_summary,
                "Job title": "",
                "First name": "",
                "Last name": "",
                "Gender": "",
                "Salutation": (make_general_salutation(ENRICH_LANGUAGE, TONE) if merged_emails else ""),  # generic if emails present
                "Emails": "; ".join(sorted(merged_emails)) if merged_emails else "",
                "Icebreaker sentence": icebreaker_sentence,
            }
            append_rows_by_headers(people_ws, header_map, [empty_dict])
            try:
                wb.save(EXCEL_PATH)
            except PermissionError:
                print(f"[excel] Save failed: file appears open/locked: {EXCEL_PATH}")
                sys.exit(1)
            after_rows = find_next_value_row(people_ws, key_cols) - 1
            print(f"[excel] {EXCEL_PATH} :: '{OUTPUT_SHEET_NAME}' — wrote company-only row; rows before={before_rows}, after={after_rows}")
    except PermissionError as pe:
        print(f"[excel] Save failed (is the file open/locked?): {pe}")
        sys.exit(1)

    return max(found_total, len(people_rows))


# =====================
# main() — bulk loop
# =====================

def main():
    # Init LLM client
    try:
        llm = LLMClient(MODEL_NAME, LLMRateLimiter(MAX_LLM_CALLS_PER_MIN))
    except Exception as e:
        print(f"[fatal] LLM init failed: {e}")
        return

    # Validate language/tone, print key config, then countdown before starting
    try:
        validate_language_and_tone(ENRICH_LANGUAGE, TONE)
    except Exception as cfg_ex:
        print(f"[fatal] Config invalid: {cfg_ex}")
        return
    print_run_config()
    countdown(10)

    try:
        # pre-check for locked Excel file (simplified: only rely on true lock detection)
        if excel_file_locked(EXCEL_PATH):
            print(f"[excel] File appears open/locked: {EXCEL_PATH}. Please close it and re-run.")
            sys.exit(1)

        wb = open_workbook()
        sites = read_input_urls(wb)  # (company, website)
        existing = read_existing_websites(wb)
        to_process = [(c, w) for (c, w) in sites if norm_url(w) not in existing]
        if MAX_ROWS_TO_PROCESS:
            to_process = to_process[:MAX_ROWS_TO_PROCESS]
        print(f"[bulk] Loaded {len(sites)} from Analyse; already enriched: {len(existing)}; processing now: {len(to_process)}")
        total_people = 0
        for i, (company, site) in enumerate(to_process, 1):
            print(f"\n[site] ({i}/{len(to_process)}) {site} — company='{company}'")
            try:
                total_people += process_site(site, company, llm, wb)
            except Exception as ex:
                print(f"[error] Unhandled for {site}: {ex}")
            finally:
                if POST_COMPANY_SLEEP_S:
                    if VERBOSE:
                        print(f"[sleep] Waiting {POST_COMPANY_SLEEP_S}s before next company...")
                    time.sleep(POST_COMPANY_SLEEP_S)
        print(f"\n[done] Total people extracted: {total_people}")
    except Exception as ex:
        print(f"[bulk] Error: {ex}")


if __name__ == '__main__':
    main()

