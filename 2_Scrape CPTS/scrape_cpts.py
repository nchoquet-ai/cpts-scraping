"""
Scraping des sites CPTS
========================
Extrait les données structurées de chaque site CPTS via Playwright + Claude.

Pipeline en cascade :
  1. Playwright charge la page équipe (rendu JS complet)
  2. Extraction texte DOM → Claude extraction texte  (90% des cas)
  3. Si DOM pauvre → Screenshot → Claude Vision       (5-10% des cas)
  4. Chiffres clés (adhérents, communes, habitants)  via rendu JS
  5. Projets / Missions                              via page dédiée ou section
  6. Actus                                           via page dédiée

Usage :
  pip install playwright pandas openpyxl anthropic tqdm
  playwright install chromium

  # Scraping (phase 1 — visite des sites)
  python scrape_cpts.py --scrape --input "Input\\CPTS_sites_verified.xlsx"

  # Test sur N lignes
  python scrape_cpts.py --scrape --input "Input\\CPTS_sites_verified.xlsx" --limit 5

  # Extraction Claude (phase 2 — analyse du contenu)
  python scrape_cpts.py --extract

  # Reprise après interruption
  python scrape_cpts.py --scrape --input "..." --resume

  # Export CSV final
  python scrape_cpts.py --export

  # Pipeline complet
  python scrape_cpts.py --all --input "Input\\CPTS_sites_verified.xlsx"

Durée estimée sur 852 CPTS :
  Phase scraping    ~2-3h  (Playwright séquentiel, ~10s/site)
  Phase extraction  ~1-2h  (Claude, 2 en parallèle)
  Total             ~4-5h
"""

import argparse
import asyncio
import base64
import json
import logging
import os
import re
import sqlite3
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

# Force UTF-8 output on Windows terminals (cp1252 can't encode emoji/box chars)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import anthropic
import pandas as pd
from tqdm import tqdm

# ─── Configuration ─────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL      = "claude-sonnet-4-6"
OUTPUT_DIR        = Path("Output")
DB_PATH           = OUTPUT_DIR / "cpts.db"
SCREENSHOTS_DIR   = OUTPUT_DIR / "screenshots"

# Playwright
PAGE_TIMEOUT      = 20_000   # ms
SCROLL_PAUSE      = 1.0      # secondes après scroll
NAV_PAUSE         = 2.0      # secondes après navigation

# Claude
CONCURRENCY_EXTRACT = 2      # appels Claude en parallèle (phase 2)
RETRY_429_WAIT      = 60     # secondes d'attente sur 429
RETRY_429_MAX       = 5      # tentatives max

# Seuil pour décider si le DOM est "pauvre" → fallback Vision
MIN_TEXT_LENGTH     = 150    # caractères minimum pour considérer le texte suffisant

# Mots-clés pages équipe (URLs)
EQUIPE_KW = [
    "equipe", "bureau", "gouvernance", "membre", "president",
    "conseil", "administration", "qui-sommes", "organisation",
    "instance", "direction", "comite", "dirigeant", "presentation",
    "association", "la-cpts", "notre-cpts", "qui-sommes-nous",
    "a-propos", "about", "structure", "gouvernance", "equipe-dirigeante",
]

# Mots-clés dans le TEXTE des pages pour détecter une page équipe
EQUIPE_TEXT_KW = [
    "président", "vice-président", "trésorier", "secrétaire",
    "bureau", "membres du bureau", "conseil d'administration",
    "gouvernance", "notre équipe", "l'équipe", "membres du ca",
    "administrateur", "co-président", "co-présidente",
    "qui sommes-nous", "qui sommes nous",
]

# Mots-clés pages projets
PROJET_KW = ["projet", "mission", "action", "programme", "initiative"]

# Mots-clés pages actus
ACTU_KW = ["actualit", "news", "agenda", "evenement", "événement", "blog"]

# Fichier de configuration des thématiques (optionnel)
THEMATIQUES_CONFIG = Path("Input/CPTS_thematiques_config.xlsx")

# ─── Chargement des thématiques ─────────────────────────────────────────────────

def load_thematiques(config_path: Path | None = None) -> dict:
    """
    Lit le fichier CPTS_thematiques_config.xlsx et retourne un dict :
    {
      'projets': [{'label': '...', 'keywords': [...], 'priorite': 1}, ...],
      'actus':   [{'label': '...', 'keywords': [...], 'priorite': 1}, ...],
    }
    Si le fichier est absent, retourne des thématiques vides (le script continue).
    """
    path = config_path or THEMATIQUES_CONFIG
    result = {"projets": [], "actus": []}

    if not path.exists():
        log.warning(f"Fichier thématiques non trouvé : {path} — mapping désactivé")
        return result

    try:
        # Onglet 1 — Thématiques Projets
        df_proj = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
        df_proj.columns = [c.strip() for c in df_proj.columns]
        for _, row in df_proj.iterrows():
            label = str(row.iloc[0]).strip()
            kw_raw = str(row.iloc[1]).strip()
            prio   = str(row.iloc[2]).strip()
            if not label or label.startswith("Thématique"):
                continue
            keywords = [k.strip().lower() for k in kw_raw.split("|") if k.strip()]
            result["projets"].append({
                "label":    label,
                "keywords": keywords,
                "priorite": int(prio) if prio.isdigit() else 3,
            })

        # Onglet 2 — Thématiques Actus
        df_actu = pd.read_excel(path, sheet_name=1, dtype=str).fillna("")
        df_actu.columns = [c.strip() for c in df_actu.columns]
        for _, row in df_actu.iterrows():
            label = str(row.iloc[0]).strip()
            kw_raw = str(row.iloc[1]).strip()
            prio   = str(row.iloc[2]).strip()
            if not label or label.startswith("Thématique"):
                continue
            keywords = [k.strip().lower() for k in kw_raw.split("|") if k.strip()]
            result["actus"].append({
                "label":    label,
                "keywords": keywords,
                "priorite": int(prio) if prio.isdigit() else 3,
            })

        log.info(f"Thématiques chargées : {len(result['projets'])} projets, {len(result['actus'])} actus")
    except Exception as e:
        log.warning(f"Erreur lecture thématiques : {e} — mapping désactivé")

    return result


def map_thematiques(text: str, thematiques: list) -> list[str]:
    """
    Mappe un texte sur les thématiques configurées.
    Retourne la liste des labels correspondants (par ordre de priorité).
    """
    if not text or not thematiques:
        return []
    text_lower = text.lower()
    matched = []
    for t in sorted(thematiques, key=lambda x: x["priorite"]):
        if any(kw in text_lower for kw in t["keywords"]):
            matched.append(t["label"])
    return matched


# Variable globale thématiques (chargée au démarrage)
THEMATIQUES: dict = {"projets": [], "actus": []}

# ─── Logging ───────────────────────────────────────────────────────────────────

def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(OUTPUT_DIR / "scrape.log", mode="a", encoding="utf-8"),
        ],
    )

log = logging.getLogger(__name__)

# ─── Base de données SQLite ─────────────────────────────────────────────────────

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    cur.executescript("""
    CREATE TABLE IF NOT EXISTS cpts (
        code            TEXT PRIMARY KEY,
        nom             TEXT,
        url             TEXT,
        scrape_status   TEXT DEFAULT 'pending',
        scrape_ts       TEXT,
        extract_status  TEXT DEFAULT 'pending',
        extract_ts      TEXT
    );

    CREATE TABLE IF NOT EXISTS pages (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        code            TEXT,
        page_type       TEXT,   -- 'equipe'|'projets'|'actus'|'home'
        url             TEXT,
        dom_text        TEXT,
        screenshot_path TEXT,
        used_vision     INTEGER DEFAULT 0,
        ts              TEXT,
        FOREIGN KEY (code) REFERENCES cpts(code)
    );

    CREATE TABLE IF NOT EXISTS extractions (
        code            TEXT PRIMARY KEY,
        equipe          TEXT,   -- JSON array
        adherents       TEXT,
        communes        TEXT,
        habitants       TEXT,
        projets         TEXT,   -- JSON array
        missions        TEXT,   -- JSON array
        actus           TEXT,   -- JSON array
        contacts        TEXT,   -- JSON object {email_regex, email_claude, telephone, adresse}
        raw_response    TEXT,
        extract_ts      TEXT,
        FOREIGN KEY (code) REFERENCES cpts(code)
    );

    CREATE TABLE IF NOT EXISTS emails (
        code            TEXT PRIMARY KEY,
        email_regex     TEXT,   -- trouvé par mailto: regex (fiable)
        email_claude    TEXT,   -- trouvé par Claude dans le texte (fallback)
        email_final     TEXT,   -- email retenu (regex prioritaire)
        source          TEXT,   -- 'regex'|'claude'|'non trouvé'
        raw_mailtos     TEXT,   -- JSON array de tous les mailto trouvés
        ts              TEXT,
        FOREIGN KEY (code) REFERENCES cpts(code)
    );
    """)
    con.commit()
    return con


def upsert_cpts(con, rows):
    cur = con.cursor()
    cur.executemany(
        "INSERT OR IGNORE INTO cpts (code, nom, url) VALUES (?,?,?)",
        rows
    )
    con.commit()


def get_pending_scrape(con, limit=None):
    cur = con.cursor()
    q = "SELECT code, nom, url FROM cpts WHERE scrape_status='pending' AND url != ''"
    if limit:
        q += f" LIMIT {limit}"
    return cur.execute(q).fetchall()


def get_pending_extract(con):
    cur = con.cursor()
    return cur.execute(
        "SELECT c.code, c.nom FROM cpts c WHERE c.scrape_status='done' AND c.extract_status='pending'"
    ).fetchall()


def save_page(con, code, page_type, url, dom_text, screenshot_path=None, used_vision=False):
    cur = con.cursor()
    cur.execute(
        """INSERT OR REPLACE INTO pages
           (code, page_type, url, dom_text, screenshot_path, used_vision, ts)
           VALUES (?,?,?,?,?,?,datetime('now'))""",
        (code, page_type, url, dom_text, screenshot_path, int(used_vision))
    )
    con.commit()


def set_scrape_status(con, code, status):
    con.execute(
        "UPDATE cpts SET scrape_status=?, scrape_ts=datetime('now') WHERE code=?",
        (status, code)
    )
    con.commit()


def set_extract_status(con, code, status):
    con.execute(
        "UPDATE cpts SET extract_status=?, extract_ts=datetime('now') WHERE code=?",
        (status, code)
    )
    con.commit()


def save_extraction(con, code, data: dict):
    cur = con.cursor()
    cur.execute(
        """INSERT OR REPLACE INTO extractions
           (code, equipe, adherents, communes, habitants, projets, missions, actus, contacts, raw_response, extract_ts)
           VALUES (?,?,?,?,?,?,?,?,?,?,datetime('now'))""",
        (
            code,
            json.dumps(data.get("equipe", []),    ensure_ascii=False),
            str(data.get("adherents", "")),
            json.dumps(data.get("communes", []),  ensure_ascii=False),
            str(data.get("habitants", "")),
            json.dumps(data.get("projets", []),   ensure_ascii=False),
            json.dumps(data.get("missions", []),  ensure_ascii=False),
            json.dumps(data.get("actus", []),     ensure_ascii=False),
            json.dumps(data.get("contacts", {}),  ensure_ascii=False),
            data.get("raw_response", ""),
        )
    )
    con.commit()

# ─── Helpers URL ───────────────────────────────────────────────────────────────

def normalize_url(raw: str) -> str:
    s = str(raw).strip().rstrip("/")
    if not s or s.lower() in {"nan", "none", "", "n/a", "non trouvé"}:
        return ""
    if s.startswith(("http://", "https://")):
        return s
    if "." in s:
        return "https://" + s
    return ""


def score_url(url: str, keywords: list) -> int:
    url_lower = url.lower()
    return sum(1 for k in keywords if k in url_lower)


def find_best_page(links: list, keywords: list) -> str | None:
    scored = [(l, score_url(l, keywords)) for l in links if score_url(l, keywords) > 0]
    scored.sort(key=lambda x: -x[1])
    return scored[0][0] if scored else None

# ─── Extraction emails ──────────────────────────────────────────────────────────

# Regex email générique (pas les images, pas les icônes)
EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE
)

# Domaines à ignorer (faux positifs fréquents)
EMAIL_BLACKLIST = {
    "sentry.io", "example.com", "test.com", "wordpress.com",
    "wixpress.com", "wix.com", "googleapis.com", "cloudflare.com",
    "schema.org", "w3.org", "jquery.com", "bootstrap.com",
}

def extract_emails_regex(html: str, dom_text: str) -> list[str]:
    """
    Approche 1 — Regex sur mailto: et texte brut.
    Fiable à ~95% quand l'email est publié.
    Priorité : mailto: > texte DOM.
    """
    found = []

    # Priorité 1 : mailto: links (le plus fiable)
    mailtos = re.findall(r'mailto:([^"\'\s>?&]+)', html, re.IGNORECASE)
    for m in mailtos:
        email = m.strip().lower().split("?")[0]  # nettoyer ?subject=...
        if EMAIL_REGEX.match(email) and not any(b in email for b in EMAIL_BLACKLIST):
            found.append(email)

    # Priorité 2 : texte DOM (si pas de mailto)
    if not found:
        for m in EMAIL_REGEX.finditer(dom_text):
            email = m.group().lower()
            if not any(b in email for b in EMAIL_BLACKLIST):
                found.append(email)

    # Dédupliquer en préservant l'ordre
    seen = set()
    result = []
    for e in found:
        if e not in seen:
            seen.add(e)
            result.append(e)

    return result[:10]  # max 10


def pick_best_email(emails: list[str], nom_cpts: str) -> str | None:
    """
    Sélectionne l'email le plus pertinent parmi une liste.
    Préfère les emails qui ressemblent à un contact CPTS.
    """
    if not emails:
        return None
    if len(emails) == 1:
        return emails[0]

    # Score : préférer les emails avec cpts/contact/info dans le nom
    def score(e):
        local = e.split("@")[0].lower()
        s = 0
        if any(k in local for k in ["cpts", "contact", "info", "secretariat", "admin"]):
            s += 3
        if any(k in local for k in ["gmail", "orange", "free", "wanadoo"]):
            s -= 1  # email perso moins probable pour contact officiel
        return s

    return sorted(emails, key=score, reverse=True)[0]


def save_email(con, code: str, email_regex: str | None,
               email_claude: str | None, raw_mailtos: list):
    """Sauvegarde les emails trouvés."""
    # Priorité : regex > claude
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



# ─── Phase 1 : Scraping Playwright ─────────────────────────────────────────────

async def get_page_content(page, url: str) -> tuple[str, str | None]:
    """
    Charge une URL, retourne (dom_text, screenshot_b64_or_None).
    Décide automatiquement si Vision est nécessaire.
    """
    try:
        # domcontentloaded puis attente du rendu JS (Wix, WordPress...)
        try:
            await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
        except Exception:
            # Retry sans timeout strict
            try:
                await page.goto(url, timeout=30000, wait_until="commit")
            except Exception as e2:
                log.warning(f"Erreur get_page_content {url}: {e2}")
                return "", None
        await asyncio.sleep(NAV_PAUSE + 1)  # +1s pour Wix/WP

        # Scroll pour déclencher le lazy-loading
        try:
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await asyncio.sleep(SCROLL_PAUSE)
            await page.evaluate("window.scrollTo(0, 0)")
            await asyncio.sleep(0.5)
        except Exception:
            pass

        # Extraction texte DOM
        dom_text = await page.evaluate("""() => {
            // Supprimer scripts, styles, nav, footer
            const remove = ['script','style','nav','footer','header','.cookie','#cookie'];
            remove.forEach(sel => {
                try { document.querySelectorAll(sel).forEach(el => el.remove()); } catch(e) {}
            });
            return document.body.innerText.replace(/\\s+/g, ' ').trim();
        }""")

        # Si texte suffisant → pas besoin de Vision
        if len(dom_text) >= MIN_TEXT_LENGTH:
            return dom_text, None

        # Texte insuffisant → screenshot pour Vision
        log.info(f"DOM pauvre ({len(dom_text)} chars) → screenshot Vision")
        screenshot_bytes = await page.screenshot(full_page=True)
        screenshot_b64   = base64.b64encode(screenshot_bytes).decode()
        return dom_text, screenshot_b64

    except Exception as e:
        log.warning(f"Erreur get_page_content {url}: {e}")
        return "", None


async def scrape_one(browser, code: str, nom: str, url: str, con) -> bool:
    """Scrape un site CPTS complet."""
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="fr-FR",
    )
    page = await context.new_page()

    try:
        # ── 1. Home page ──────────────────────────────────────────────────────
        log.info(f"[{code}] Scraping {url}")
        await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
        await asyncio.sleep(NAV_PAUSE)
        # Attendre que le body soit chargé
        try:
            await page.wait_for_selector("body", timeout=5000)
        except Exception:
            pass

        # Récupérer tous les liens internes
        base_domain = urlparse(url).netloc
        all_links = await page.evaluate(f"""() => {{
            const domain = '{base_domain}';
            return [...new Set(
                Array.from(document.querySelectorAll('a[href]'))
                    .map(a => a.href)
                    .filter(h => h.includes(domain) && !h.includes('#')
                                 && !h.match(/\\.(pdf|jpg|png|css|js)$/i))
            )].slice(0, 80);
        }}""")

        # Texte + chiffres de la home
        home_text, home_ss = await get_page_content(page, url)
        save_page(con, code, "home", url, home_text,
                  screenshot_path=None, used_vision=False)

        # ── Email — Approche 1 : regex mailto sur le HTML brut ────────────────
        home_html = await page.content()
        raw_mailtos = re.findall(r'mailto:([^"\'\s>?&]+)', home_html, re.IGNORECASE)
        raw_mailtos = [m.strip().lower().split("?")[0] for m in raw_mailtos]
        email_regex = pick_best_email(
            extract_emails_regex(home_html, home_text), nom
        )

        # Chercher aussi sur la page contact si elle existe
        contact_url = find_best_page(all_links, ["contact", "nous-contacter", "nous-ecrire"])
        if contact_url and not email_regex:
            try:
                await page.goto(contact_url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                await asyncio.sleep(1)
                contact_html = await page.content()
                contact_text = await page.evaluate("() => document.body.innerText")
                email_regex = pick_best_email(
                    extract_emails_regex(contact_html, contact_text), nom
                )
                raw_mailtos += re.findall(r'mailto:([^"\'\s>?&]+)', contact_html, re.IGNORECASE)
                # Revenir à la home pour la suite
                await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                await asyncio.sleep(1)
            except Exception as e:
                log.debug(f"[{code}] Page contact inaccessible: {e}")

        # Stocker le résultat regex (Claude complétera en phase 2 si vide)
        save_email(con, code, email_regex, None, list(set(raw_mailtos)))
        log.info(f"[{code}] Email regex: {email_regex or '(non trouvé)'}")

        # ── 2. Page équipe ────────────────────────────────────────────────────
        # Étape 1 : score URL
        equipe_url = find_best_page(all_links, EQUIPE_KW)

        # Étape 2 : si aucune URL ne matche, on visite les pages courtes du menu
        # et on cherche les mots-clés bureau dans leur texte
        if not equipe_url:
            # Pages candidates = liens courts (pas d'articles, pas de /20xx/)
            candidates = [l for l in all_links
                         if l != url
                         and not any(x in l for x in ["pdf", "jpg", "png", "wp-content",
                                                       "mentions", "confidential", "login",
                                                       "inscription", "adherere", "don"])
                         and len(l.replace(url, "").strip("/").split("/")) <= 2][:15]
            for candidate in candidates:
                try:
                    await page.goto(candidate, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                    await asyncio.sleep(1)
                    cand_text = await page.evaluate("() => document.body.innerText")
                    cand_lower = cand_text.lower()
                    if any(kw in cand_lower for kw in EQUIPE_TEXT_KW):
                        equipe_url = candidate
                        log.info(f"[{code}] Page équipe trouvée par texte : {candidate}")
                        break
                except Exception:
                    continue
            # Revenir à la home si on a navigué
            if equipe_url:
                await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                await asyncio.sleep(1)

        equipe_url = equipe_url or url
        log.info(f"[{code}] Page équipe retenue : {equipe_url}")
        if equipe_url != url:
            equipe_text, equipe_ss = await get_page_content(page, equipe_url)
        else:
            log.warning(f"[{code}] ⚠️ Fallback home — page équipe non trouvée")
            equipe_text, equipe_ss = home_text, home_ss

        # Sauvegarder screenshot si Vision nécessaire
        ss_path = None
        if equipe_ss:
            ss_path = str(SCREENSHOTS_DIR / f"{code}_equipe.png")
            with open(ss_path, "wb") as f:
                f.write(base64.b64decode(equipe_ss))

        save_page(con, code, "equipe", equipe_url, equipe_text,
                  screenshot_path=ss_path, used_vision=bool(equipe_ss))

        # ── 3. Page projets ───────────────────────────────────────────────────
        projets_url = find_best_page(all_links, PROJET_KW)
        if projets_url and projets_url != equipe_url:
            projets_text, _ = await get_page_content(page, projets_url)
            save_page(con, code, "projets", projets_url, projets_text)

        # ── 4. Page actus (titres seulement, pas le détail) ───────────────────
        actu_url = find_best_page(all_links, ACTU_KW)
        if actu_url:
            actu_text, _ = await get_page_content(page, actu_url)
            save_page(con, code, "actus", actu_url, actu_text)

        set_scrape_status(con, code, "done")
        log.info(f"[{code}] ✅ Scraping OK")
        return True

    except Exception as e:
        log.error(f"[{code}] ❌ Erreur scraping: {e}")
        set_scrape_status(con, code, "error")
        return False
    finally:
        await context.close()


async def run_scraping(con, limit=None, resume=False):
    """Lance le scraping Playwright sur toutes les CPTS pending."""
    from playwright.async_api import async_playwright

    if not resume:
        # Remettre les erreurs en pending pour retry
        con.execute("UPDATE cpts SET scrape_status='pending' WHERE scrape_status='error'")
        con.commit()

    todo = get_pending_scrape(con, limit)
    log.info(f"Phase 1 — {len(todo)} CPTS à scraper")

    if not todo:
        log.info("Rien à scraper.")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        for code, nom, url in tqdm(todo, desc="Scraping"):
            await scrape_one(browser, code, nom, url, con)
            await asyncio.sleep(1)  # politesse
        await browser.close()

    log.info("Phase 1 terminée.")

# ─── Phase 2 : Extraction Claude ────────────────────────────────────────────────

EXTRACT_PROMPT = """\
Tu es un expert en santé publique française. Analyse le contenu de ce site CPTS et extrais les informations structurées.

CPTS : {nom}

=== CONTENU PAGE ÉQUIPE ===
{equipe_text}

=== CONTENU PAGE PROJETS/MISSIONS ===
{projets_text}

=== CONTENU HOME (chiffres clés) ===
{home_text}

=== CONTENU ACTUS ===
{actu_text}

Réponds UNIQUEMENT avec ce JSON (sans markdown) :
{{
  "equipe": [
    {{"nom": "...", "prenom": "...", "civilite": "M.|Mme|Dr", "fonction": "Président|Vice-Président|Trésorier|Secrétaire|...", "specialite": "Médecin généraliste|Infirmier|...|null"}}
  ],
  "adherents": "nombre ou null",
  "communes": ["liste", "des", "communes"],
  "habitants": "nombre ou null",
  "projets": [
    {{"titre": "...", "description": "...", "mots_cles": ["prévention", "accès aux soins", "..."]}}
  ],
  "missions": ["mission 1", "mission 2"],
  "actus": [
    {{"titre": "...", "date": "JJ/MM/AAAA ou null", "resume": "..."}}
  ],
  "contacts": {{"email": "...", "telephone": "...", "adresse": "..."}}
}}

Règles :
- equipe : extrais TOUS les membres du bureau/CA visibles (président, vice-président, trésorier, secrétaire, administrateurs)
  * Cherche dans toutes les sections : "Les membres du bureau", "Notre équipe", "Gouvernance", "Bureau", "L'équipe", "Conseil d'administration"
  * Format du texte souvent : "Fonction\nPrénom Nom" ou "Prénom Nom\nFonction" ou "M./Mme/Dr Prénom NOM - Fonction"
  * Si civilité absente, déduis-la du prénom (Marie → Mme, Jean → M.)
  * Si spécialité absente, laisse null (ne pas inventer)
  * Inclus les administrateurs même sans spécialité mentionnée
- adherents : cherche "X adhérents", "X professionnels", "X membres", "X soignants" — prends le nombre le plus récent
- projets : uniquement les projets concrets avec un titre identifiable
- actus : les 5 plus récentes uniquement
- mots_cles_projets : choisis UNIQUEMENT parmi les thématiques listées ci-dessous (plusieurs possibles)
- mots_cles_actus : idem, uniquement parmi les thématiques listées ci-dessous
- Si une info est absente, mets null ou liste vide []
- Ne jamais inventer de noms ou fonctions

=== THÉMATIQUES AUTORISÉES POUR LES PROJETS ===
{thematiques_projets}

=== THÉMATIQUES AUTORISÉES POUR LES ACTUS ===
{thematiques_actus}
"""

VISION_PROMPT = """\
Tu es un expert en santé publique française. Cette image montre la page équipe/bureau d'une CPTS.

Extrais tous les membres visibles avec leurs informations.

Réponds UNIQUEMENT avec ce JSON (sans markdown) :
{{
  "equipe": [
    {{"nom": "...", "prenom": "...", "civilite": "M.|Mme|Dr", "fonction": "...", "specialite": "..."}}
  ]
}}

Si tu ne peux pas lire clairement un nom ou une fonction, omets ce membre.
"""


def _call_claude_text(client, nom, pages: dict) -> dict:
    """Appel Claude extraction texte avec retry 429."""
    equipe_text  = pages.get("equipe",  "")[:6000]
    projets_text = pages.get("projets", "")[:3000]
    home_text    = pages.get("home",    "")[:4000]  # 4000 pour capturer les chiffres en bas de page
    actu_text    = pages.get("actus",   "")[:2000]

    # Formater les thématiques pour le prompt
    them_proj = THEMATIQUES.get("projets", [])
    them_actu = THEMATIQUES.get("actus",   [])
    them_proj_str = "\n".join(
        f"- {t['label']} (mots-clés : {', '.join(t['keywords'][:5])})"
        for t in them_proj
    ) or "(aucune thématique configurée — utilise des mots-clés libres)"
    them_actu_str = "\n".join(
        f"- {t['label']} (mots-clés : {', '.join(t['keywords'][:5])})"
        for t in them_actu
    ) or "(aucune thématique configurée — utilise des mots-clés libres)"

    prompt = EXTRACT_PROMPT.format(
        nom=nom,
        equipe_text=equipe_text or "(non disponible)",
        projets_text=projets_text or "(non disponible)",
        home_text=home_text or "(non disponible)",
        actu_text=actu_text or "(non disponible)",
        thematiques_projets=them_proj_str,
        thematiques_actus=them_actu_str,
    )

    for attempt in range(RETRY_429_MAX + 1):
        try:
            resp = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}],
            )
            text = resp.content[0].text.strip()
            return _parse_json(text)
        except Exception as e:
            if "429" in str(e) and attempt < RETRY_429_MAX:
                log.warning(f"429 extraction — attente {RETRY_429_WAIT}s")
                time.sleep(RETRY_429_WAIT)
                continue
            log.error(f"Erreur extraction texte {nom}: {e}")
            return {}
    return {}


def _call_claude_vision(client, screenshot_path: str) -> dict:
    """Appel Claude Vision sur un screenshot."""
    try:
        with open(screenshot_path, "rb") as f:
            img_data = base64.b64encode(f.read()).decode()

        resp = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1000,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_data}},
                    {"type": "text",  "text": VISION_PROMPT},
                ]
            }],
        )
        text = resp.content[0].text.strip()
        return _parse_json(text)
    except Exception as e:
        log.error(f"Erreur Vision {screenshot_path}: {e}")
        return {}


def _parse_json(text: str) -> dict:
    """Parsing JSON robuste."""
    if not text:
        return {}
    # Direct
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # Nettoyer backticks
    clean = text.strip()
    if clean.startswith("```"):
        clean = clean.split("\n", 1)[-1] if "\n" in clean else clean[3:]
    if clean.endswith("```"):
        clean = clean.rsplit("```", 1)[0]
    try:
        return json.loads(clean.strip())
    except json.JSONDecodeError:
        pass
    # Extraire premier {...}
    start = text.find("{")
    end   = text.rfind("}")
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end+1])
        except json.JSONDecodeError:
            pass
    return {}



def _extract_email_claude(client, nom: str, text: str) -> str | None:
    """
    Approche 2 — Claude cherche un email de contact dans le texte.
    Fallback si le regex n'a rien trouvé (email en image, obfusqué, etc.)
    """
    if not text.strip():
        return None

    prompt = f"""Voici le contenu textuel du site de la CPTS "{nom}".
Cherche l'adresse email de contact officielle de cette CPTS (pas les emails personnels des membres).

Contenu :
{text[:3000]}

Réponds UNIQUEMENT avec l'adresse email si tu en trouves une, ou "non trouvé" si tu n'en vois pas.
Ne donne aucune explication. Exemple de réponse : contact@cpts-exemple.fr"""

    for attempt in range(RETRY_429_MAX + 1):
        try:
            resp = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=50,
                messages=[{"role": "user", "content": prompt}],
            )
            result = resp.content[0].text.strip().lower()
            if "non trouvé" in result or not EMAIL_REGEX.match(result):
                return None
            return result
        except Exception as e:
            if "429" in str(e) and attempt < RETRY_429_MAX:
                time.sleep(RETRY_429_WAIT)
                continue
            return None
    return None

def extract_one(client, code: str, nom: str) -> bool:
    """Extrait les données d'une CPTS depuis la DB (ouvre sa propre connexion — thread-safe)."""
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    # Récupérer les pages scrapées
    rows = cur.execute(
        "SELECT page_type, dom_text, screenshot_path, used_vision FROM pages WHERE code=?",
        (code,)
    ).fetchall()

    pages = {r[0]: r[1] for r in rows}
    equipe_row = next((r for r in rows if r[0] == "equipe"), None)

    result = {}

    # ── Extraction texte principale ──────────────────────────────────────────
    result = _call_claude_text(client, nom, pages)
    result["raw_response"] = json.dumps(result, ensure_ascii=False)

    # ── Fallback Vision si équipe vide ────────────────────────────────────────
    if not result.get("equipe") and equipe_row and equipe_row[2]:
        log.info(f"[{code}] Équipe vide → fallback Vision")
        vision_result = _call_claude_vision(client, equipe_row[2])
        if vision_result.get("equipe"):
            result["equipe"] = vision_result["equipe"]

    # ── Email — Approche 2 : fallback Claude si regex n'a rien trouvé ────────
    email_row = cur.execute(
        "SELECT email_regex, email_claude FROM emails WHERE code=?", (code,)
    ).fetchone()

    email_regex  = email_row[0] if email_row else None
    email_claude = email_row[1] if email_row else None

    if not email_regex:
        # Claude cherche dans tout le texte disponible
        all_text = " ".join(filter(None, [
            pages.get("home", ""),
            pages.get("equipe", ""),
            pages.get("projets", ""),
        ]))[:4000]

        email_claude = _extract_email_claude(client, nom, all_text)

        if email_claude:
            log.info(f"[{code}] Email Claude (fallback): {email_claude}")
        else:
            log.info(f"[{code}] Email non trouvé (regex + Claude)")

        # Mise à jour de la table emails avec le résultat Claude
        save_email(con, code, email_regex, email_claude,
                   json.loads(cur.execute(
                       "SELECT raw_mailtos FROM emails WHERE code=?", (code,)
                   ).fetchone()[0] if cur.execute(
                       "SELECT raw_mailtos FROM emails WHERE code=?", (code,)
                   ).fetchone() else "[]"))
    else:
        log.info(f"[{code}] Email regex déjà trouvé: {email_regex}")

    # Enrichir contacts avec l'email final
    email_final = email_regex or email_claude
    contacts = result.get("contacts", {})
    if email_final and not contacts.get("email"):
        contacts["email"] = email_final
    contacts["email_regex"]  = email_regex  or ""
    contacts["email_claude"] = email_claude or ""
    contacts["email_source"] = "regex" if email_regex else ("claude" if email_claude else "non trouvé")
    result["contacts"] = contacts

    save_extraction(con, code, result)
    set_extract_status(con, code, "done")
    con.close()
    log.info(f"[{code}] ✅ Extraction OK — {len(result.get('equipe', []))} membres")
    return True


async def run_extraction(con):
    """Lance l'extraction Claude sur toutes les CPTS scrapées."""
    if not ANTHROPIC_API_KEY:
        log.error("ANTHROPIC_API_KEY manquante.")
        return

    todo = get_pending_extract(con)
    log.info(f"Phase 2 — {len(todo)} CPTS à extraire")
    if not todo:
        log.info("Rien à extraire.")
        return

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    sem    = asyncio.Semaphore(CONCURRENCY_EXTRACT)

    async def extract_bounded(code, nom):
        async with sem:
            return await asyncio.get_event_loop().run_in_executor(
                None, extract_one, client, code, nom
            )

    tasks = [extract_bounded(code, nom) for code, nom in todo]
    for coro in tqdm(asyncio.as_completed(tasks), total=len(tasks), desc="Extraction"):
        await coro

    log.info("Phase 2 terminée.")

# ─── Export CSV ─────────────────────────────────────────────────────────────────

def run_export(con):
    """Génère les CSV de sortie depuis la DB."""
    cur = con.cursor()

    # ── CSV équipe dirigeante ────────────────────────────────────────────────
    rows_equipe = []
    for code, nom, url in cur.execute("SELECT code, nom, url FROM cpts").fetchall():
        ext = cur.execute("SELECT equipe, adherents, communes, habitants, contacts FROM extractions WHERE code=?", (code,)).fetchone()
        if not ext:
            continue
        equipe   = json.loads(ext[0] or "[]")
        contacts = json.loads(ext[4] or "{}")
        # Email final depuis la table emails
        email_row = cur.execute(
            "SELECT email_final, email_regex, email_claude, source FROM emails WHERE code=?",
            (code,)
        ).fetchone()
        email_final  = email_row[0] if email_row else contacts.get("email", "")
        email_regex  = email_row[1] if email_row else ""
        email_claude = email_row[2] if email_row else ""
        email_source = email_row[3] if email_row else ""

        for m in equipe:
            rows_equipe.append({
                "Code CPTS":        code,
                "Nom CPTS":         nom,
                "URL":              url,
                "Civilité":         m.get("civilite", ""),
                "Prénom":           m.get("prenom", ""),
                "Nom":              m.get("nom", ""),
                "Fonction":         m.get("fonction", ""),
                "Spécialité":       m.get("specialite", ""),
                "Email CPTS":       email_final or "",
                "Email (regex)":    email_regex or "",
                "Email (Claude)":   email_claude or "",
                "Source email":     email_source or "",
                "Tél CPTS":         contacts.get("telephone", ""),
            })

    # ── CSV chiffres clés ────────────────────────────────────────────────────
    rows_chiffres = []
    for code, nom, url in cur.execute("SELECT code, nom, url FROM cpts").fetchall():
        ext = cur.execute("SELECT adherents, communes, habitants FROM extractions WHERE code=?", (code,)).fetchone()
        if not ext:
            continue
        communes = json.loads(ext[1] or "[]")
        email_row_c = cur.execute(
            "SELECT email_final, source FROM emails WHERE code=?", (code,)
        ).fetchone()
        contacts = json.loads(cur.execute(
            "SELECT contacts FROM extractions WHERE code=?", (code,)
        ).fetchone()[0] or "{}")
        rows_chiffres.append({
            "Code CPTS":      code,
            "Nom CPTS":       nom,
            "URL":            url,
            "Email contact":  email_row_c[0] if email_row_c else "",
            "Source email":   email_row_c[1] if email_row_c else "",
            "Téléphone":      contacts.get("telephone", ""),
            "Adhérents":      ext[0] or "",
        })

    # ── CSV projets ──────────────────────────────────────────────────────────
    rows_projets = []
    them_proj = THEMATIQUES.get("projets", [])
    them_actu = THEMATIQUES.get("actus",   [])
    # Colonnes dynamiques une par thématique (1/0)
    proj_labels = [t["label"] for t in them_proj]
    actu_labels = [t["label"] for t in them_actu]

    for code, nom in cur.execute("SELECT code, nom FROM cpts").fetchall():
        ext = cur.execute("SELECT projets, missions, actus FROM extractions WHERE code=?", (code,)).fetchone()
        if not ext:
            continue
        projets  = json.loads(ext[0] or "[]")
        missions = json.loads(ext[1] or "[]")
        for p in projets:
            mots_cles_claude = p.get("mots_cles", [])
            texte_projet = p.get("titre","") + " " + p.get("description","")
            # Mapping thématiques : Claude + regex sur le texte
            mapped = set(mots_cles_claude)
            mapped.update(map_thematiques(texte_projet, them_proj))
            row = {
                "Code CPTS":   code,
                "Nom CPTS":    nom,
                "Type":        "Projet",
                "Titre":       p.get("titre", ""),
                "Description": p.get("description", ""),
                "Thématiques": ", ".join(sorted(mapped)),
            }
            # Colonnes 0/1 par thématique
            for lbl in proj_labels:
                row[lbl] = 1 if lbl in mapped else 0
            rows_projets.append(row)
        for m in missions:
            mapped = set(map_thematiques(m, them_proj))
            row = {
                "Code CPTS":   code,
                "Nom CPTS":    nom,
                "Type":        "Mission",
                "Titre":       m,
                "Description": "",
                "Thématiques": ", ".join(sorted(mapped)),
            }
            for lbl in proj_labels:
                row[lbl] = 1 if lbl in mapped else 0
            rows_projets.append(row)

    # ── CSV actus ─────────────────────────────────────────────────────────────
    rows_actus = []
    for code, nom in cur.execute("SELECT code, nom FROM cpts").fetchall():
        ext = cur.execute("SELECT actus FROM extractions WHERE code=?", (code,)).fetchone()
        if not ext:
            continue
        actus = json.loads(ext[0] or "[]")
        for a in actus:
            texte_actu = a.get("titre","") + " " + a.get("resume","")
            mapped = set(map_thematiques(texte_actu, them_actu))
            mapped.update(a.get("mots_cles", []))
            row = {
                "Code CPTS":   code,
                "Nom CPTS":    nom,
                "Titre":       a.get("titre", ""),
                "Date":        a.get("date", ""),
                "Résumé":      a.get("resume", ""),
                "Thématiques": ", ".join(sorted(mapped)),
            }
            for lbl in actu_labels:
                row[lbl] = 1 if lbl in mapped else 0
            rows_actus.append(row)

    # ── Sauvegarde ───────────────────────────────────────────────────────────
    pd.DataFrame(rows_equipe).to_csv(OUTPUT_DIR / "cpts_equipe.csv",    index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_chiffres).to_csv(OUTPUT_DIR / "cpts_chiffres.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_projets).to_csv(OUTPUT_DIR / "cpts_projets.csv",  index=False, encoding="utf-8-sig")
    pd.DataFrame(rows_actus).to_csv(OUTPUT_DIR / "cpts_actus.csv",      index=False, encoding="utf-8-sig")

    log.info(f"Export terminé :")
    log.info(f"  {len(rows_equipe)} membres équipe      → cpts_equipe.csv")
    log.info(f"  {len(rows_chiffres)} CPTS chiffres     → cpts_chiffres.csv")
    log.info(f"  {len(rows_projets)} projets/missions   → cpts_projets.csv")
    log.info(f"  {len(rows_actus)} actus                → cpts_actus.csv")
    log.info(f"  Thématiques projets mappées : {len(proj_labels)}")
    log.info(f"  Thématiques actus mappées   : {len(actu_labels)}")

# ─── Statistiques ───────────────────────────────────────────────────────────────

def print_stats(con):
    cur = con.cursor()
    total    = cur.execute("SELECT COUNT(*) FROM cpts").fetchone()[0]
    scraped  = cur.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='done'").fetchone()[0]
    errors   = cur.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='error'").fetchone()[0]
    extracted= cur.execute("SELECT COUNT(*) FROM cpts WHERE extract_status='done'").fetchone()[0]
    vision   = cur.execute("SELECT COUNT(*) FROM pages WHERE used_vision=1").fetchone()[0]

    print("\n─── Résumé ─────────────────────────────────────────────")
    print(f"  Total CPTS          : {total}")
    print(f"  ✅ Scrapées         : {scraped}")
    print(f"  ❌ Erreurs scraping : {errors}")
    print(f"  🔍 Extraites Claude : {extracted}")
    print(f"  🖼️  Vision utilisée  : {vision} pages")
    print("────────────────────────────────────────────────────────\n")

# ─── Pipeline principal ─────────────────────────────────────────────────────────

async def run(args):
    global THEMATIQUES
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    con = init_db()

    # Charger les thématiques depuis le fichier de config
    config_path = Path(args.thematiques) if args.thematiques else THEMATIQUES_CONFIG
    THEMATIQUES = load_thematiques(config_path)

    # Charger le fichier Excel si scraping demandé
    if args.scrape or args.all:
        if not args.input:
            log.error("--input requis pour le scraping.")
            return

        df = pd.read_excel(args.input, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]

        # Colonnes flexibles
        url_col  = next((c for c in df.columns if any(k in c.lower() for k in ["site", "url", "lien"])), None)
        nom_col  = next((c for c in df.columns if any(k in c.lower() for k in ["nom", "label", "libellé"])), None)
        code_col = next((c for c in df.columns if any(k in c.lower() for k in ["code", "id"])), df.columns[0])

        if not url_col or not nom_col:
            log.error(f"Colonnes non trouvées. Colonnes disponibles : {df.columns.tolist()}")
            return

        log.info(f"Colonnes détectées → code: '{code_col}' | nom: '{nom_col}' | url: '{url_col}'"  )

        rows = []
        for _, row in df.iterrows():
            url = normalize_url(row.get(url_col, ""))
            if url:
                rows.append((row[code_col], row[nom_col], url))

        upsert_cpts(con, rows)
        log.info(f"{len(rows)} CPTS chargées (avec URL valide)")

        if args.limit:
            # Marquer les CPTS hors limite comme skip
            all_codes = [r[0] for r in rows]
            keep = set(all_codes[:args.limit])
            con.execute(
                f"UPDATE cpts SET scrape_status='skip' WHERE code NOT IN ({','.join('?'*len(keep))})",
                list(keep)
            )
            con.commit()

        await run_scraping(con, limit=args.limit, resume=args.resume)

    if args.extract or args.all:
        await run_extraction(con)

    if args.export or args.all:
        run_export(con)

    print_stats(con)
    con.close()

# ─── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Scraping des sites CPTS",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input",   help="Fichier Excel source (CPTS_sites_verified.xlsx)")
    parser.add_argument("--limit",   type=int, help="Limiter aux N premières CPTS")
    parser.add_argument("--resume",  action="store_true", help="Reprendre sans reset des erreurs")
    parser.add_argument("--scrape",  action="store_true", help="Phase 1 : scraping Playwright")
    parser.add_argument("--extract", action="store_true", help="Phase 2 : extraction Claude")
    parser.add_argument("--export",  action="store_true", help="Export CSV final")
    parser.add_argument("--all",     action="store_true", help="Pipeline complet (scrape + extract + export)")
    parser.add_argument("--thematiques", default=None, help="Fichier Excel thématiques (défaut: Input/CPTS_thematiques_config.xlsx)")
    parser.add_argument("--stats",   action="store_true", help="Afficher les stats de la DB")
    args = parser.parse_args()

    if not any([args.scrape, args.extract, args.export, args.all, args.stats]):
        parser.print_help()
    else:
        setup_logging()
        if args.stats:
            con = init_db()
            print_stats(con)
            con.close()
        else:
            asyncio.run(run(args))
