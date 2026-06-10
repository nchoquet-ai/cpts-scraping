"""
Scraping CPTS — PHASE 2 : Analyse & extraction (Claude)
========================================================
Lit les pages brutes capturées par scrape_phase1.py (table `pages`, immuable),
extrait les données structurées via Claude, mappe sur le référentiel multi-axes
(A Mission / B Pathologies CIM-10 / C Public / D Vie CPTS / E Vulnérabilité),
puis exporte en CSV.

Principes :
  - PAS de routage : on agrège le texte de TOUTES les pages → 1 appel holistique/CPTS
    (chunké si volumineux). Claude trie le contenu lui-même.
  - Vision ciblée équipe + pages à DOM pauvre (sélection par contenu, plafonnée).
  - Provenance : chaque objet porte source_url (page d'origine).
  - Pathologies : Claude détecte en texte libre → codage CIM-10 en cascade
    (table curée → fichier officiel 2054 → hors-référentiel).
  - La table `pages` n'est JAMAIS modifiée → Phase 2 relançable à volonté.

Usage :
  set ANTHROPIC_API_KEY=sk-ant-...
  python scrape_phase2.py --extract --export
  python scrape_phase2.py --extract --limit 11
  python scrape_phase2.py --export
  python scrape_phase2.py --stats
"""

import argparse
import asyncio
import base64
import csv
import difflib
import json
import logging
import os
import re
import sqlite3
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import urlparse

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import anthropic
import pandas as pd
from tqdm import tqdm

# ─── Configuration ──────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL      = "claude-sonnet-4-6"
OUTPUT_DIR        = Path("Output")
DB_PATH           = OUTPUT_DIR / "cpts.db"

CATEGORIES_CONFIG = Path("Input/CPTS_categories_config.xlsx")
CIM10_REFERENCE   = Path("Input/CIM10_FR2025_niveau3.csv")

CONCURRENCY_EXTRACT = 4
RETRY_429_WAIT      = 60
RETRY_429_MAX       = 5

CHUNK_CHARS    = 150_000   # taille max d'un appel holistique (≈ 37k tokens)
MIN_TEXT_LEN   = 80        # DOM "pauvre" → candidat Vision
MIN_CONTENT    = 200       # pages "coquilles" (< N chars) exclues de l'appel TEXTE Claude
                           # (réduit les tokens ; elles restent candidates Vision)
VISION_MAX     = 4         # nb max de screenshots envoyés à Vision par CPTS
FUZZY_THRESHOLD = 0.62     # seuil de rapprochement CIM-10 (étape b)

# Mots-clés équipe pour sélectionner les pages à passer en Vision (par CONTENU)
EQUIPE_VISION_KW = [
    "président", "présidente", "vice-président", "co-président", "trésorier",
    "secrétaire", "bureau", "organigramme", "trombinoscope", "notre équipe",
    "conseil d'administration", "membres du bureau", "gouvernance", "le bureau",
]
# Indices d'URL d'une page "équipe" (organigramme-image souvent sans mots-clés texte)
# → sert UNIQUEMENT à sélectionner les screenshots à passer en Vision (pas au routage).
EQUIPE_URL_KW = [
    "equipe", "association", "bureau", "gouvernance", "qui-sommes", "conseil",
    "administration", "membre", "organigramme", "trombinoscope", "la-cpts",
    "notre-cpts", "presentation", "a-propos", "instance", "direction",
]

# ─── Logging ──────────────────────────────────────────────────────────────────

def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(OUTPUT_DIR / "scrape_phase2.log", mode="a", encoding="utf-8"),
        ],
    )

log = logging.getLogger(__name__)

# ─── Base de données ──────────────────────────────────────────────────────────

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    # Ajout colonnes extract_* sur cpts (Phase 1 ne les a pas créées)
    for ddl in (
        "ALTER TABLE cpts ADD COLUMN extract_status TEXT DEFAULT 'pending'",
        "ALTER TABLE cpts ADD COLUMN extract_ts TEXT",
    ):
        try:
            cur.execute(ddl)
        except sqlite3.OperationalError:
            pass  # colonne déjà présente

    cur.executescript("""
    CREATE TABLE IF NOT EXISTS extractions (
        code         TEXT PRIMARY KEY,
        equipe       TEXT,   -- JSON []
        adherents    TEXT,   -- JSON {nombre, source_url}
        projets      TEXT,   -- JSON [] (projets/missions)
        actus        TEXT,   -- JSON []
        contacts     TEXT,   -- JSON {}
        raw_response TEXT,
        extract_ts   TEXT
    );
    CREATE TABLE IF NOT EXISTS emails (
        code         TEXT PRIMARY KEY,
        email_regex  TEXT,
        email_claude TEXT,
        email_final  TEXT,
        source       TEXT,
        raw_mailtos  TEXT,
        ts           TEXT
    );
    CREATE TABLE IF NOT EXISTS pathologies (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        code          TEXT,   -- code CPTS
        objet_type    TEXT,   -- 'projet' | 'actu'
        objet_titre   TEXT,
        pathologie_texte TEXT,
        code_cim10    TEXT,
        libelle_officiel TEXT,
        source_code   TEXT,   -- 'curé' | 'auto-2054' | 'hors-réf'
        a_verifier    INTEGER,
        source_url    TEXT,
        ts            TEXT
    );
    """)
    con.commit()
    return con


def get_pending_extract(con):
    return con.execute(
        "SELECT code, nom FROM cpts WHERE scrape_status='done' "
        "AND (extract_status IS NULL OR extract_status='pending')"
    ).fetchall()


def set_extract_status(con, code, status):
    con.execute(
        "UPDATE cpts SET extract_status=?, extract_ts=datetime('now') WHERE code=?",
        (status, code)
    )
    con.commit()


def save_extraction(con, code, data: dict):
    con.execute(
        """INSERT OR REPLACE INTO extractions
           (code, equipe, adherents, projets, actus, contacts, raw_response, extract_ts)
           VALUES (?,?,?,?,?,?,?,datetime('now'))""",
        (
            code,
            json.dumps(data.get("equipe", []),    ensure_ascii=False),
            json.dumps(data.get("adherents", {}), ensure_ascii=False),
            json.dumps(data.get("projets", []),   ensure_ascii=False),
            json.dumps(data.get("actus", []),     ensure_ascii=False),
            json.dumps(data.get("contacts", {}),  ensure_ascii=False),
            data.get("raw_response", ""),
        )
    )
    con.commit()


def save_pathologies(con, code, rows):
    con.execute("DELETE FROM pathologies WHERE code=?", (code,))
    con.executemany(
        """INSERT INTO pathologies
           (code, objet_type, objet_titre, pathologie_texte, code_cim10,
            libelle_officiel, source_code, a_verifier, source_url, ts)
           VALUES (?,?,?,?,?,?,?,?,?,datetime('now'))""",
        rows
    )
    con.commit()

# ─── Emails (réutilisé de scrape_cpts.py) ─────────────────────────────────────

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", re.IGNORECASE)
EMAIL_BLACKLIST = {
    "sentry.io", "example.com", "test.com", "wordpress.com", "wixpress.com",
    "wix.com", "googleapis.com", "cloudflare.com", "schema.org", "w3.org",
    "jquery.com", "bootstrap.com",
}

def extract_emails_regex(html: str, dom_text: str) -> list[str]:
    found = []
    for m in re.findall(r'mailto:([^"\'\s>?&]+)', html or "", re.IGNORECASE):
        email = m.strip().lower().split("?")[0]
        if EMAIL_REGEX.match(email) and not any(b in email for b in EMAIL_BLACKLIST):
            found.append(email)
    if not found:
        for m in EMAIL_REGEX.finditer(dom_text or ""):
            email = m.group().lower()
            if not any(b in email for b in EMAIL_BLACKLIST):
                found.append(email)
    seen, res = set(), []
    for e in found:
        if e not in seen:
            seen.add(e); res.append(e)
    return res[:10]

def pick_best_email(emails: list[str]) -> str | None:
    if not emails:
        return None
    if len(emails) == 1:
        return emails[0]
    def score(e):
        local = e.split("@")[0].lower()
        s = 0
        if any(k in local for k in ["cpts", "contact", "info", "secretariat", "admin"]):
            s += 3
        if any(k in local for k in ["gmail", "orange", "free", "wanadoo"]):
            s -= 1
        return s
    return sorted(emails, key=score, reverse=True)[0]

def save_email(con, code, email_regex, email_claude, raw_mailtos):
    if email_regex:
        email_final, source = email_regex, "regex"
    elif email_claude:
        email_final, source = email_claude, "claude"
    else:
        email_final, source = None, "non trouvé"
    con.execute(
        """INSERT OR REPLACE INTO emails
           (code, email_regex, email_claude, email_final, source, raw_mailtos, ts)
           VALUES (?,?,?,?,?,?,datetime('now'))""",
        (code, email_regex, email_claude, email_final, source,
         json.dumps(raw_mailtos, ensure_ascii=False))
    )
    con.commit()

# ─── Référentiel de catégories ────────────────────────────────────────────────

# Globaux chargés au démarrage
CATEGORIES: dict = {}      # {axe: [{code,label,keywords[],cim10,libelle_off}]}
CIM10_LABELS: dict = {}    # {code3: libellé officiel}
CIM10_LIST: list = []      # [(libellé_lower, code3, libellé)]

AXES_KW = ["A", "C", "D", "E"]  # axes mappés par mots-clés (B = cascade pathologie)

def _norm(s: str) -> str:
    """Minuscule + suppression des accents (matching robuste è/é, etc.)."""
    s = (s or "").lower()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _kw_match(kw: str, text_norm: str) -> bool:
    """
    Mot-clé présent dans le texte normalisé.
    - mots-clés ≤ 3 car. (ag, sas, hpv, ald…) : MOT ENTIER (évite les sous-chaînes)
    - mots-clés ≥ 4 car. : PRÉFIXE de mot (gère pluriels « projet→projets » et
      stems « diabèt→diabète/diabétique »)
    """
    k = _norm(kw)
    if not k:
        return False
    if len(k.replace(" ", "")) <= 3:
        pat = r"(?<!\w)" + re.escape(k) + r"(?!\w)"
    else:
        pat = r"(?<!\w)" + re.escape(k) + r"\w*"
    return re.search(pat, text_norm) is not None

def load_categories(path: Path) -> dict:
    """Lit le .xlsx 5 onglets → {axe: [ {code,label,keywords,cim10,libelle_off} ]}."""
    result = {}
    sheets = {
        "A": "A_Mission", "B": "B_Pathologies", "C": "C_Public",
        "D": "D_VieCPTS", "E": "E_Vulnerabilite",
    }
    for axe, sheet in sheets.items():
        df = pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")
        items = []
        for _, r in df.iterrows():
            kw = [k.strip().lower() for k in str(r.get("Mots-clés", "")).split("|") if k.strip()]
            items.append({
                "code":        str(r.get("Code", "")).strip(),
                "label":       str(r.get("Catégorie", "")).strip(),
                "keywords":    kw,
                "cim10":       str(r.get("CIM-10", "")).strip(),
                "libelle_off": str(r.get("Libellé CIM-10 officiel", "")).strip(),
            })
        result[axe] = items
    log.info("Référentiel chargé : " + ", ".join(f"{a}={len(result[a])}" for a in result))
    return result

def load_cim10(path: Path):
    labels, lst = {}, []
    if not path.exists():
        log.warning(f"Fichier CIM-10 absent : {path} — étape (b) désactivée")
        return labels, lst
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f, delimiter=";"):
            code, lib = row["code"], row["libelle"]
            labels[code] = lib
            lst.append((_norm(lib), code, lib))   # clé normalisée (accents retirés)
    log.info(f"CIM-10 officiel chargé : {len(labels)} codes")
    return labels, lst

def map_axes_keywords(text: str) -> dict:
    """Mappe un texte sur les axes A/C/D/E par mots-clés. → {axe: [(code,label)]}."""
    tn = _norm(text)
    out = {}
    for axe in AXES_KW:
        hits = []
        for cat in CATEGORIES.get(axe, []):
            if any(_kw_match(kw, tn) for kw in cat["keywords"]):
                hits.append((cat["code"], cat["label"]))
        out[axe] = hits
    return out

def code_pathologie(text: str) -> dict:
    """
    Cascade de codage CIM-10 d'une pathologie en texte libre.
    a) table curée B  b) rapprochement /2054 (avec garde-fou token)  c) hors-référentiel
    """
    raw = (text or "").strip()
    tn = _norm(raw)
    if not tn:
        return {}
    # a) table curée (mots-clés, frontière de mot, accents normalisés)
    for cat in CATEGORIES.get("B", []):
        if cat["keywords"] and any(_kw_match(kw, tn) for kw in cat["keywords"]):
            return {"code_cim10": cat["cim10"], "libelle_officiel": cat["libelle_off"] or cat["label"],
                    "categorie_b": cat["label"], "source_code": "curé", "a_verifier": 0}
    # b) rapprochement sur les libellés officiels (2054) — avec garde-fou : il faut
    #    partager un token significatif (≥4 lettres), sinon on rejette (évite les
    #    faux rapprochements type "endométriose" → "rétention d'urine").
    if CIM10_LIST:
        words_t = {w for w in re.findall(r"\w{4,}", tn)}
        best_ratio, best = 0.0, None
        for lib_low, code, lib in CIM10_LIST:
            r = difflib.SequenceMatcher(None, tn, lib_low).ratio()
            if r > best_ratio:
                words_l = set(re.findall(r"\w{4,}", lib_low))
                if words_t & words_l:           # garde-fou : token partagé obligatoire
                    best_ratio, best = r, (code, lib)
        if best and best_ratio >= FUZZY_THRESHOLD:
            return {"code_cim10": best[0], "libelle_officiel": best[1],
                    "categorie_b": "", "source_code": "auto-2054", "a_verifier": 1}
    # c) hors référentiel
    return {"code_cim10": "", "libelle_officiel": "", "categorie_b": "",
            "source_code": "hors-réf", "a_verifier": 1}

# ─── Prompts Claude ───────────────────────────────────────────────────────────

EXTRACT_PROMPT = """\
Tu es un expert en santé publique française. Analyse le contenu agrégé d'un site CPTS
(plusieurs pages, chacune précédée de son URL) et extrais les informations structurées.

CPTS : {nom}

=== CONTENU DES PAGES (chaque bloc commence par --- PAGE (url) ---) ===
{pages_content}

Réponds UNIQUEMENT avec ce JSON (sans markdown), en remplissant "source_url" avec l'URL
de la page d'où provient chaque information :
{{
  "equipe": [
    {{"civilite":"M.|Mme|Dr|null","prenom":"...","nom":"...","fonction":"Président|Vice-Président|Trésorier|Secrétaire|Administrateur|Coordinateur|...","specialite":"Médecin généraliste|Infirmier|...|null","source_url":"..."}}
  ],
  "adherents": {{"nombre":"nombre le plus récent ou null","source_url":"..."}},
  "contacts": {{"email":"...|null","telephone":"...|null","adresse":"...|null","source_url":"..."}},
  "projets": [
    {{"titre":"...","description":"...","date":"AAAA ou JJ/MM/AAAA ou null","pathologies":["pathologie en clair", "..."],"source_url":"..."}}
  ],
  "actus": [
    {{"titre":"...","date":"JJ/MM/AAAA ou null","resume":"...","pathologies":["pathologie en clair"],"source_url":"..."}}
  ]
}}

Règles :
- equipe : TOUS les membres du bureau/CA visibles. Déduis la civilité du prénom si absente. specialite=null si non mentionnée. Ne jamais inventer.
- adherents : le nombre de professionnels/adhérents de la CPTS, valeur la plus récente.
- projets : projets ET missions de la CPTS (même catégorie). Titre + description concrète.
- pathologies : pour chaque projet/actu, liste les pathologies/maladies/thèmes de santé
  mentionnés, EN CLAIR et au plus près du texte (ex. "insuffisance cardiaque", "diabète",
  "cancer du sein", "santé mentale"). Liste vide si aucune pathologie précise.
- actus : les actualités/événements. Résumé court.
- source_url : TOUJOURS l'URL exacte de la page (bloc --- PAGE) où l'info a été trouvée.
- Si une info est absente : null ou liste vide [].
"""

VISION_PROMPT = """\
Tu es un expert en santé publique française. Cette image montre une page (équipe/bureau ou
organigramme) d'une CPTS. Extrais tous les membres visibles.

Réponds UNIQUEMENT avec ce JSON (sans markdown) :
{
  "equipe": [
    {"civilite":"M.|Mme|Dr|null","prenom":"...","nom":"...","fonction":"...","specialite":"...|null"}
  ]
}
Extrais tous les membres identifiables, même partiellement. N'omets un membre que si son nom est totalement illisible.
"""

def _parse_json(text: str) -> dict:
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    clean = text.strip()
    if clean.startswith("```"):
        clean = clean.split("\n", 1)[-1] if "\n" in clean else clean[3:]
    if clean.endswith("```"):
        clean = clean.rsplit("```", 1)[0]
    try:
        return json.loads(clean.strip())
    except json.JSONDecodeError:
        pass
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e > s:
        try:
            return json.loads(text[s:e+1])
        except json.JSONDecodeError:
            pass
    return {}

def _claude_call(client, messages, max_tokens):
    for attempt in range(RETRY_429_MAX + 1):
        try:
            resp = client.messages.create(model=CLAUDE_MODEL, max_tokens=max_tokens, messages=messages)
            return resp.content[0].text.strip()
        except Exception as ex:
            if "429" in str(ex) and attempt < RETRY_429_MAX:
                log.warning(f"429 — attente {RETRY_429_WAIT}s")
                time.sleep(RETRY_429_WAIT)
                continue
            log.error(f"Erreur appel Claude : {ex}")
            return ""
    return ""

def _call_claude_text(client, nom, pages_content) -> dict:
    prompt = EXTRACT_PROMPT.format(nom=nom, pages_content=pages_content)
    return _parse_json(_claude_call(client, [{"role": "user", "content": prompt}], 16000))

def _image_b64(ss_path, max_dim=7500) -> str | None:
    """Charge un screenshot et le redimensionne si une dimension dépasse max_dim
    (limite Claude = 8000 px). Retourne le base64 PNG, ou None si illisible."""
    try:
        from PIL import Image
        import io
        img = Image.open(ss_path)
        w, h = img.size
        if max(w, h) > max_dim:
            scale = max_dim / max(w, h)
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))))
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode()
    except Exception as ex:
        log.warning(f"Screenshot illisible/redim impossible {ss_path}: {ex}")
        return None

def _call_claude_vision(client, ss_path) -> dict:
    img = _image_b64(ss_path)
    if not img:
        return {}
    return _parse_json(_claude_call(client, [{
        "role": "user",
        "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img}},
            {"type": "text", "text": VISION_PROMPT},
        ],
    }], 1500))

def _extract_email_claude(client, nom, text) -> str | None:
    if not text.strip():
        return None
    prompt = (f'Voici le contenu du site de la CPTS "{nom}". Donne UNIQUEMENT son adresse '
              f'email de contact officielle, ou "non trouvé".\n\n{text[:3000]}')
    res = _claude_call(client, [{"role": "user", "content": prompt}], 50).lower()
    if not res or "non trouvé" in res or not EMAIL_REGEX.match(res):
        return None
    return res

# ─── Extraction d'une CPTS ────────────────────────────────────────────────────

def _build_chunks(rows):
    """rows = [(url, dom_text)]. Regroupe en chunks <= CHUNK_CHARS, blocs étiquetés par URL."""
    chunks, cur, size = [], [], 0
    for url, txt in rows:
        block = f"--- PAGE ({url}) ---\n{(txt or '').strip()}\n"
        if size + len(block) > CHUNK_CHARS and cur:
            chunks.append("\n".join(cur)); cur, size = [], 0
        cur.append(block); size += len(block)
    if cur:
        chunks.append("\n".join(cur))
    return chunks

def _merge_text_results(results: list[dict]) -> dict:
    """Fusionne les résultats de plusieurs chunks."""
    merged = {"equipe": [], "projets": [], "actus": [], "adherents": {}, "contacts": {}}
    seen_membres, seen_proj, seen_actu = set(), set(), set()
    for r in results:
        for m in r.get("equipe", []) or []:
            key = ((m.get("prenom") or "") + (m.get("nom") or "")).lower().strip()
            if key and key not in seen_membres:
                seen_membres.add(key); merged["equipe"].append(m)
        for p in r.get("projets", []) or []:
            key = (p.get("titre") or "").lower().strip()
            if key and key not in seen_proj:
                seen_proj.add(key); merged["projets"].append(p)
        for a in r.get("actus", []) or []:
            key = (a.get("titre") or "").lower().strip()
            if key and key not in seen_actu:
                seen_actu.add(key); merged["actus"].append(a)
        adh = r.get("adherents") or {}
        if adh.get("nombre") and not merged["adherents"].get("nombre"):
            merged["adherents"] = adh
        ct = r.get("contacts") or {}
        for k in ("email", "telephone", "adresse", "source_url"):
            if ct.get(k) and not merged["contacts"].get(k):
                merged["contacts"][k] = ct[k]
    return merged

def extract_one(client, code, nom) -> bool:
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    pages = cur.execute(
        "SELECT url, dom_text, screenshot_path, page_type FROM pages WHERE code=?", (code,)
    ).fetchall()
    if not pages:
        log.warning(f"[{code}] aucune page — skip")
        set_extract_status(con, code, "done"); con.close(); return False

    # ── 1. Extraction holistique texte (chunké) ──────────────────────────────
    # Filtre les pages "coquilles" (< MIN_CONTENT chars) pour économiser les tokens.
    text_rows = [(u, t) for u, t, _, _ in pages if len((t or "").strip()) >= MIN_CONTENT]
    if not text_rows:  # site entièrement image → garder ce qu'on a
        text_rows = [(u, t) for u, t, _, _ in pages]
    n_skip = len(pages) - len(text_rows)
    if n_skip:
        log.info(f"[{code}] {n_skip} pages coquilles exclues du texte (< {MIN_CONTENT} chars)")
    chunks = _build_chunks(text_rows)
    results = [_call_claude_text(client, nom, ch) for ch in chunks]
    result = _merge_text_results(results)

    # ── 2. Vision ciblée équipe + DOM pauvre (sélection par contenu) ──────────
    #     Priorité : organigramme natif (-1) > URL équipe (0) > mots-clés (1) > DOM pauvre (2)
    vision_candidates = []
    for url, txt, ss, ptype in pages:
        if not ss:
            continue
        tl = (txt or "").lower()
        path = urlparse(url).path.lower()
        is_org    = (ptype == "organigramme")                    # image organigramme native
        is_url    = any(k in path for k in EQUIPE_URL_KW)        # URL "équipe" (organigramme)
        is_equipe = any(kw in tl for kw in EQUIPE_VISION_KW)     # mots-clés équipe dans le texte
        is_poor   = len(tl.strip()) < MIN_TEXT_LEN               # page image (DOM pauvre)
        if is_org or is_url or is_equipe or is_poor:
            prio = -1 if is_org else (0 if is_url else (1 if is_equipe else 2))
            vision_candidates.append((prio, ss, url))
    vision_candidates.sort(key=lambda x: x[0])
    for _, ss, url in vision_candidates[:VISION_MAX]:
        vres = _call_claude_vision(client, ss)
        existing = {((m.get("prenom") or "") + (m.get("nom") or "")).lower() for m in result["equipe"]}
        for m in vres.get("equipe", []) or []:
            key = ((m.get("prenom") or "") + (m.get("nom") or "")).lower()
            if key and key not in existing:
                m.setdefault("source_url", url)
                result["equipe"].append(m); existing.add(key)

    # ── 3. Email : regex sur toutes les pages → fallback Claude ───────────────
    home_html_text = " ".join((t or "") for _, t, _, _ in pages)
    emails = extract_emails_regex("", home_html_text)
    email_regex = pick_best_email(emails)
    email_claude = None
    if not email_regex:
        email_claude = _extract_email_claude(client, nom, home_html_text[:4000])
    save_email(con, code, email_regex, email_claude, emails)
    email_final = email_regex or email_claude
    ct = result.get("contacts") or {}
    if email_final and not ct.get("email"):
        ct["email"] = email_final
    result["contacts"] = ct

    # ── 4. Codage pathologies (cascade) + mapping axes → enrichit projets/actus ─
    patho_rows = []
    for objet_type, items in (("projet", result.get("projets", [])),
                              ("actu",   result.get("actus", []))):
        for it in items:
            texte = ((it.get("titre") or "") + " " + (it.get("description") or "") + " " + (it.get("resume") or "")).strip()
            axes = map_axes_keywords(texte)
            it["axes"] = axes
            codes_b = []
            for pt in (it.get("pathologies") or []):
                pc = code_pathologie(pt)
                if not pc:
                    continue
                codes_b.append({"pathologie": pt, **pc})
                patho_rows.append((code, objet_type, it.get("titre", ""), pt,
                                   pc["code_cim10"], pc["libelle_officiel"],
                                   pc["source_code"], pc["a_verifier"], it.get("source_url", "")))
            it["cim10"] = codes_b
    save_pathologies(con, code, patho_rows)

    result["raw_response"] = json.dumps({"chunks": len(chunks)}, ensure_ascii=False)
    save_extraction(con, code, result)
    set_extract_status(con, code, "done")
    con.close()
    log.info(f"[{code}] ✅ {len(result['equipe'])} membres | "
             f"{len(result['projets'])} projets | {len(result['actus'])} actus | "
             f"{len(patho_rows)} pathologies")
    return True

# ─── Orchestration ────────────────────────────────────────────────────────────

async def run_extraction(con, limit=None):
    if not ANTHROPIC_API_KEY:
        log.error("ANTHROPIC_API_KEY manquante.")
        return
    todo = get_pending_extract(con)
    if limit:
        todo = todo[:limit]
    log.info(f"Phase 2 — {len(todo)} CPTS à extraire (concurrence={CONCURRENCY_EXTRACT})")
    if not todo:
        log.info("Rien à extraire."); return

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    sem = asyncio.Semaphore(CONCURRENCY_EXTRACT)
    done = {"n": 0}

    async def bounded(code, nom):
        async with sem:
            await asyncio.get_event_loop().run_in_executor(None, extract_one, client, code, nom)
            done["n"] += 1
            if done["n"] % 25 == 0:
                log.info(f"Checkpoint {done['n']}/{len(todo)}")

    await asyncio.gather(*(bounded(c, n) for c, n in todo))
    log.info("Phase 2 terminée.")

# ─── Export CSV ───────────────────────────────────────────────────────────────

def run_export(con):
    cur = con.cursor()
    proj_cats = {a: [c["label"] for c in CATEGORIES.get(a, [])] for a in ["A", "B", "C", "D", "E"]}

    rows_eq, rows_adh, rows_proj, rows_actu = [], [], [], []

    for code, nom, url in cur.execute("SELECT code, nom, url FROM cpts").fetchall():
        ext = cur.execute(
            "SELECT equipe, adherents, projets, actus, contacts FROM extractions WHERE code=?", (code,)
        ).fetchone()
        if not ext:
            continue
        equipe   = json.loads(ext[0] or "[]")
        adher    = json.loads(ext[1] or "{}")
        projets  = json.loads(ext[2] or "[]")
        actus    = json.loads(ext[3] or "[]")
        contacts = json.loads(ext[4] or "{}")
        erow = cur.execute("SELECT email_final, source FROM emails WHERE code=?", (code,)).fetchone()
        email_final = erow[0] if erow else contacts.get("email", "")

        for m in equipe:
            rows_eq.append({
                "Code CPTS": code, "Nom CPTS": nom,
                "Civilité": m.get("civilite", ""), "Prénom": m.get("prenom", ""),
                "Nom": m.get("nom", ""), "Fonction": m.get("fonction", ""),
                "Spécialité": m.get("specialite", ""),
                "Email CPTS": email_final or "", "Source URL": m.get("source_url", ""),
            })

        rows_adh.append({
            "Code CPTS": code, "Nom CPTS": nom, "URL": url,
            "Adhérents": adher.get("nombre", ""), "Source URL": adher.get("source_url", ""),
            "Email contact": email_final or "", "Téléphone": contacts.get("telephone", ""),
            "Adresse": contacts.get("adresse", ""),
        })

        def cat_row(it, objet_type, rows_target):
            axes = it.get("axes", {})
            cim = it.get("cim10", [])
            base = {
                "Code CPTS": code, "Nom CPTS": nom, "Type": objet_type,
                "Titre": it.get("titre", ""),
                "Description": it.get("description", "") or it.get("resume", ""),
                "Date": it.get("date", ""), "Source URL": it.get("source_url", ""),
                "Pathologies (texte)": " | ".join(it.get("pathologies", []) or []),
                "CIM-10": " | ".join(c["code_cim10"] for c in cim if c.get("code_cim10")),
                "CIM-10 libellés": " | ".join(c["libelle_officiel"] for c in cim if c.get("libelle_officiel")),
                "À vérifier": 1 if any(c.get("a_verifier") for c in cim) else 0,
            }
            # colonnes 0/1 par catégorie de chaque axe
            matched_b = {c.get("categorie_b") for c in cim if c.get("categorie_b")}
            for axe in ["A", "C", "D", "E"]:
                labels = {lbl for _, lbl in axes.get(axe, [])}
                for lbl in proj_cats[axe]:
                    base[f"{axe}:{lbl}"] = 1 if lbl in labels else 0
            for lbl in proj_cats["B"]:
                base[f"B:{lbl}"] = 1 if lbl in matched_b else 0
            rows_target.append(base)

        for p in projets:
            cat_row(p, "Projet/Mission", rows_proj)
        for a in actus:
            cat_row(a, "Actu", rows_actu)

    pd.DataFrame(rows_eq).to_csv(OUTPUT_DIR / "cpts_equipe.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_adh).to_csv(OUTPUT_DIR / "cpts_adherents.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_proj).to_csv(OUTPUT_DIR / "cpts_projets.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_actu).to_csv(OUTPUT_DIR / "cpts_actus.csv", index=False, encoding="utf-8-sig")

    log.info("Export terminé :")
    log.info(f"  {len(rows_eq)} membres        → cpts_equipe.csv")
    log.info(f"  {len(rows_adh)} CPTS adhérents → cpts_adherents.csv")
    log.info(f"  {len(rows_proj)} projets/missions → cpts_projets.csv")
    log.info(f"  {len(rows_actu)} actus          → cpts_actus.csv")


def report_enrichissement(con):
    """Rapport des pathologies auto-2054 + hors-réf, dédupliquées, par fréquence."""
    rows = con.execute(
        "SELECT pathologie_texte, code_cim10, libelle_officiel, source_code, COUNT(DISTINCT code) "
        "FROM pathologies WHERE source_code != 'curé' "
        "GROUP BY lower(pathologie_texte) ORDER BY 5 DESC"
    ).fetchall()
    if not rows:
        return
    out = OUTPUT_DIR / "pathologies_a_enrichir.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["pathologie_texte", "code_cim10_propose", "libelle_officiel", "source", "nb_cpts"])
        w.writerows(rows)
    log.info(f"  {len(rows)} pathologies à enrichir → {out.name}")

# ─── Stats ────────────────────────────────────────────────────────────────────

def print_stats(con):
    cur = con.cursor()
    total = cur.execute("SELECT COUNT(*) FROM cpts").fetchone()[0]
    scraped = cur.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='done'").fetchone()[0]
    try:
        extr = cur.execute("SELECT COUNT(*) FROM cpts WHERE extract_status='done'").fetchone()[0]
    except sqlite3.OperationalError:
        extr = 0
    print("\n─── Résumé Phase 2 ─────────────────────────────────────")
    print(f"  Total CPTS     : {total}")
    print(f"  Scrapées (P1)  : {scraped}")
    print(f"  Extraites (P2) : {extr}")
    print("────────────────────────────────────────────────────────\n")

# ─── Pipeline ─────────────────────────────────────────────────────────────────

async def run(args):
    global CATEGORIES, CIM10_LABELS, CIM10_LIST
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    con = init_db()
    CATEGORIES = load_categories(Path(args.categories) if args.categories else CATEGORIES_CONFIG)
    CIM10_LABELS, CIM10_LIST = load_cim10(CIM10_REFERENCE)

    if args.extract:
        await run_extraction(con, limit=args.limit)
    if args.export:
        run_export(con)
        report_enrichissement(con)
    print_stats(con)
    con.close()

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Scraping CPTS — Phase 2 (analyse Claude)")
    p.add_argument("--extract", action="store_true", help="Extraction Claude")
    p.add_argument("--export",  action="store_true", help="Export CSV")
    p.add_argument("--limit",   type=int, help="Limiter aux N premières CPTS")
    p.add_argument("--categories", default=None, help="Fichier de config catégories (.xlsx)")
    p.add_argument("--concurrency", type=int, default=CONCURRENCY_EXTRACT, help="Appels parallèles")
    p.add_argument("--stats", action="store_true", help="Stats DB")
    args = p.parse_args()

    CONCURRENCY_EXTRACT = args.concurrency
    if not any([args.extract, args.export, args.stats]):
        p.print_help()
    else:
        setup_logging()
        if args.stats and not (args.extract or args.export):
            c = init_db(); print_stats(c); c.close()
        else:
            asyncio.run(run(args))
