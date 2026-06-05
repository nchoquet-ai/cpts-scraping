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
import httpx
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
SCROLL_PAUSE      = 2.0      # secondes après scroll (lazy-loading WordPress)
NAV_PAUSE         = 2.0      # secondes après navigation

# Claude
CONCURRENCY_EXTRACT = 2      # appels Claude en parallèle (phase 2)
RETRY_429_WAIT      = 60     # secondes d'attente sur 429
RETRY_429_MAX       = 5      # tentatives max

# Seuil pour décider si le DOM est "pauvre" → fallback Vision
MIN_TEXT_LENGTH     = 80     # caractères minimum pour considérer le texte suffisant
MAX_PAGES_TO_SCRAPE = 20    # pages candidates visitées pour détection textuelle équipe

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
    "membres du bureau", "conseil d'administration",
    "notre équipe", "l'équipe dirigeante",
    "président", "trésorier", "secrétaire",
    "membres du ca", "administrateur", "co-président",
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
    # Ignorer les URLs d'articles : le DERNIER segment du chemin a trop de tirets
    # Ex. article : /la-cpts-des-3-provinces-forme-les-professionnels-… (14 tirets)
    # Ex. valide  : /qui-sommes-nous/conseil-d-administration-et-bureau  (4 tirets)
    last_segment = urlparse(url).path.strip("/").split("/")[-1]
    if len(last_segment.split("-")) > 6:  # slug trop long = article de blog
        return 0
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
            await asyncio.sleep(2.0)  # attente supplémentaire pour animations JS
        except Exception:
            pass

        # Extraction texte DOM
        dom_text = await page.evaluate("""() => {
            // Supprimer scripts, styles, nav, footer
            const remove = ['script','style','nav','footer','header','.cookie','#cookie'];
            remove.forEach(sel => {
                try { document.querySelectorAll(sel).forEach(el => el.remove()); } catch(e) {}
            });
            // Dédupliquer les figcaptions (Elementor/Wix les duplique souvent)
            const captions = document.querySelectorAll('figcaption, [class*=caption]');
            const captionTexts = [...new Set(Array.from(captions).map(el => el.textContent.trim()).filter(t => t.length > 0))];
            const seenCaptions = new Set();
            Array.from(captions).forEach(el => {
                const t = el.textContent.trim();
                if (seenCaptions.has(t)) { el.remove(); } else { seenCaptions.add(t); }
            });
            return document.body.innerText.replace(/\\s+/g, ' ').trim();
        }""")

        screenshot_b64 = None

        if len(dom_text) < MIN_TEXT_LENGTH:
            # DOM pauvre (Wix/WP image-based) → scroll pour forcer rendu + screenshot Vision
            try:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await asyncio.sleep(1.5)
                await page.evaluate("window.scrollTo(0, 0)")
                await asyncio.sleep(0.5)
            except Exception:
                pass
            log.info(f"DOM pauvre ({len(dom_text)} chars) → screenshot Vision")
            screenshot_bytes = await page.screenshot(full_page=True)
            screenshot_b64   = base64.b64encode(screenshot_bytes).decode()
        else:
            # DOM suffisant → détecter organigramme image et forcer screenshot si trouvé
            try:
                page_html = await page.content()
                if re.search(r'organigramme|org.?chart|trombinoscope', page_html, re.IGNORECASE):
                    # Essayer d'abord de récupérer l'image organigramme directement
                    org_imgs = await page.evaluate("""() =>
                        Array.from(document.querySelectorAll('img'))
                            .filter(i => i.alt && /organigramme|org.?chart|trombinoscope/i.test(i.alt))
                            .map(i => ({src: i.src, alt: i.alt}))
                    """)
                    if org_imgs:
                        try:
                            img_url = org_imgs[0]['src']
                            # Wix sert souvent en AVIF malgré l'extension PNG
                            # Forcer le PNG natif en remplaçant les paramètres de transco Wix
                            if 'wixstatic.com' in img_url:
                                import re as re2
                                img_url = re2.sub(r'enc_avif[^/]*', 'enc_png', img_url)
                                img_url = re2.sub(r'quality_auto[^/]*', '', img_url)
                            r = httpx.get(img_url, timeout=15, follow_redirects=True)
                            if r.status_code == 200:
                                # Vérifier le vrai Content-Type
                                ct = r.headers.get('content-type', 'image/png')
                                if 'avif' in ct:
                                    # Re-essayer sans les paramètres Wix
                                    base_url_img = img_url.split('/v1/')[0] + '/v1/fill/w_1200/' + img_url.split('/')[-1]
                                    r2 = httpx.get(base_url_img, timeout=15, follow_redirects=True)
                                    if r2.status_code == 200:
                                        r = r2
                                screenshot_b64 = base64.b64encode(r.content).decode()
                                log.info(f"Image organigramme récupérée ({r.headers.get('content-type','?')}): {org_imgs[0]['alt']}")
                        except Exception as e_img:
                            log.debug(f"Erreur récupération image organigramme : {e_img}")

                    if not screenshot_b64:
                        # Fallback : screenshot pleine page
                        log.info("Organigramme détecté → screenshot Vision forcé")
                        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                        await asyncio.sleep(1.5)
                        await page.evaluate("window.scrollTo(0, 0)")
                        await asyncio.sleep(0.5)
                        screenshot_bytes = await page.screenshot(full_page=True)
                        screenshot_b64   = base64.b64encode(screenshot_bytes).decode()
            except Exception:
                pass

        return dom_text, screenshot_b64

    except Exception as e:
        log.warning(f"Erreur get_page_content {url}: {e}")
        return "", None


async def discover_and_scrape_pages(
    page, home_url: str, all_links: list, code: str = ""
) -> dict:
    """
    Découvre et scrape les pages candidates (équipe, projets, actus).
    Retourne : {url: {"text": str, "screenshot_b64": str|None, "page_type": str}}
    La home n'est PAS incluse (déjà sauvegardée par scrape_one).
    """
    result = {}
    prefix = f"[{code}] " if code else ""

    # ── 1. Page équipe ─────────────────────────────────────────────────────────
    equipe_url = find_best_page(all_links, EQUIPE_KW)

    # Fallback texte : visiter les candidats et détecter via EQUIPE_TEXT_KW
    if not equipe_url:
        def priority(link: str) -> int:
            score = score_url(link, EQUIPE_KW + PROJET_KW)
            if "cpts" in urlparse(link).path.lower():
                score += 4
            return score

        candidates = sorted(all_links, key=priority, reverse=True)[:MAX_PAGES_TO_SCRAPE]
        for cand_url in candidates:
            try:
                await page.goto(cand_url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                await asyncio.sleep(1)
                cand_text = await page.evaluate(
                    "() => document.body.innerText.replace(/\\s+/g, ' ').trim()"
                )
                text_lower = cand_text.lower()
                if sum(1 for kw in EQUIPE_TEXT_KW if kw in text_lower) >= 2:
                    log.info(f"{prefix}Page équipe trouvée par texte : {cand_url}")
                    equipe_url = cand_url
                    break
            except Exception:
                pass

    if equipe_url:
        log.info(f"{prefix}Page équipe retenue : {equipe_url}")
    else:
        equipe_url = home_url
        log.warning(f"{prefix}⚠️ Fallback home — page équipe non trouvée")
        log.info(f"{prefix}Page équipe retenue : {equipe_url}")

    equipe_text, equipe_ss = await get_page_content(page, equipe_url)
    result[equipe_url] = {
        "text": equipe_text,
        "screenshot_b64": equipe_ss,
        "page_type": "equipe",
    }

    # ── 2. Page projets ────────────────────────────────────────────────────────
    projets_url = find_best_page(all_links, PROJET_KW)
    if projets_url and projets_url not in result:
        projets_text, _ = await get_page_content(page, projets_url)
        result[projets_url] = {
            "text": projets_text,
            "screenshot_b64": None,
            "page_type": "projets",
        }

    # ── 3. Page actus ──────────────────────────────────────────────────────────
    actu_url = find_best_page(all_links, ACTU_KW)
    if actu_url and actu_url not in result:
        actu_text, _ = await get_page_content(page, actu_url)
        result[actu_url] = {
            "text": actu_text,
            "screenshot_b64": None,
            "page_type": "actus",
        }

    return result


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

        # ── Home page : texte + email ─────────────────────────────────────
        home_text, _ = await get_page_content(page, url)
        save_page(con, code, "home", url, home_text,
                  screenshot_path=None, used_vision=False)

        home_html = await page.content()
        raw_mailtos = re.findall(r'mailto:([^"\'\s>?&]+)', home_html, re.IGNORECASE)
        raw_mailtos = [m.strip().lower().split("?")[0] for m in raw_mailtos]
        email_regex = pick_best_email(extract_emails_regex(home_html, home_text), nom)
        save_email(con, code, email_regex, None, list(set(raw_mailtos)))
        log.info(f"[{code}] Email home: {email_regex or '(non trouvé)'}")

        # ── Option C : Discovery + scraping Content-First ────────────────────
        log.info(f"[{code}] Discovery des pages candidates...")
        scraped_pages = await discover_and_scrape_pages(page, url, all_links, code)

        vision_count = 0
        for page_url, data in scraped_pages.items():
            ss_path = None
            if data["screenshot_b64"]:
                vision_count += 1
                ss_path = str(SCREENSHOTS_DIR / f"{code}_{vision_count}.png")
                with open(ss_path, "wb") as f:
                    f.write(base64.b64decode(data["screenshot_b64"]))
            save_page(con, code, data["page_type"], page_url,
                      data["text"], screenshot_path=ss_path,
                      used_vision=bool(data["screenshot_b64"]))
            # Email dans chaque page visitée
            if not email_regex:
                try:
                    await page.goto(page_url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
                    ph = await page.content()
                    email_regex = pick_best_email(
                        extract_emails_regex(ph, data["text"]), nom
                    )
                    raw_mailtos += re.findall(r'mailto:([^"\'\s>?&]+)', ph, re.IGNORECASE)
                    if email_regex:
                        save_email(con, code, email_regex, None, list(set(raw_mailtos)))
                        log.info(f"[{code}] Email trouvé sur {page_url}: {email_regex}")
                except Exception:
                    pass

        log.info(f"[{code}] {len(scraped_pages)} pages scrapées | {vision_count} screenshots")

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
Pages analysées : {nb_pages}

=== CONTENU DES PAGES DU SITE ===
{pages_content}

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

Extrais tous les membres que tu peux identifier, même partiellement. En cas de doute sur l'orthographe, inclus ta meilleure lecture. N'omets un membre que si son nom est totalement illisible.
"""


def _call_claude_text(client, nom, pages_rows: list) -> dict:
    """
    Option C — Appel Claude avec toutes les pages agrégées.
    pages_rows : liste de (page_type, url, dom_text, screenshot_path)
    """
    # Formater les thématiques
    them_proj = THEMATIQUES.get("projets", [])
    them_actu = THEMATIQUES.get("actus",   [])
    them_proj_str = "\n".join(
        f"- {t['label']} (mots-clés : {', '.join(t['keywords'][:5])})"
        for t in them_proj
    ) or "(aucune thématique configurée)"
    them_actu_str = "\n".join(
        f"- {t['label']} (mots-clés : {', '.join(t['keywords'][:5])})"
        for t in them_actu
    ) or "(aucune thématique configurée)"

    # Construire le contenu agrégé de toutes les pages
    # Budget tokens : ~12000 chars total répartis entre les pages
    pages_content_parts = []
    
    # Home en priorité (chiffres clés)
    home_rows = [r for r in pages_rows if r[0] == "home"]
    other_rows = [r for r in pages_rows if r[0] != "home"]
    
    char_budget = 16000

    for ptype, purl, ptext, _ in (home_rows + other_rows):
        if not ptext or not ptext.strip():
            continue
        # Budget par page : plus pour equipe (membres souvent en fin de page)
        if ptype == "equipe":
            limit = 6000
        elif ptype == "home":
            limit = 3000
        elif ptype == "other":
            limit = 2500
        else:
            limit = 1500
        
        truncated = ptext.strip()[:limit]
        pages_content_parts.append(
            f"--- PAGE : {ptype.upper()} ({purl}) ---\n{truncated}"
        )
        char_budget -= len(truncated)
        if char_budget <= 0:
            break

    pages_content = "\n\n".join(pages_content_parts) if pages_content_parts else "(aucun contenu disponible)"

    prompt = EXTRACT_PROMPT.format(
        nom=nom,
        nb_pages=len(pages_content_parts),
        pages_content=pages_content,
        thematiques_projets=them_proj_str,
        thematiques_actus=them_actu_str,
    )

    for attempt in range(RETRY_429_MAX + 1):
        try:
            resp = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=2500,
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
    """
    Option C — Extrait les données d'une CPTS.
    Passe TOUTES les pages à Claude en une seule fois.
    Fallback Vision sur les pages avec screenshot si équipe vide.
    """
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    # Récupérer toutes les pages scrapées (type, url, texte, screenshot)
    rows = cur.execute(
        "SELECT page_type, url, dom_text, screenshot_path FROM pages WHERE code=?",
        (code,)
    ).fetchall()

    # Dict page_type → dom_text (pour recherche email fallback)
    pages = {ptype: (ptext or "") for ptype, _, ptext, _ in rows}

    result = {}

    # ── Extraction texte — toutes les pages agrégées ─────────────────────────
    result = _call_claude_text(client, nom, rows)
    result["raw_response"] = json.dumps(result, ensure_ascii=False)

    # ── Vision systématique : fusionner TOUS les screenshots avec le texte ──────
    screenshot_rows = [r for r in rows if r[3]]  # pages avec screenshot_path
    for ptype, purl, _, ss_path in screenshot_rows:
        log.info(f"[{code}] Vision sur {ptype} ({purl})")
        vision_result = _call_claude_vision(client, ss_path)
        if vision_result.get("equipe"):
            existing_noms = {m.get("nom", "").lower() for m in result.get("equipe", [])}
            for membre in vision_result["equipe"]:
                if membre.get("nom", "").lower() not in existing_noms:
                    result.setdefault("equipe", []).append(membre)
                    existing_noms.add(membre.get("nom", "").lower())
            log.info(f"[{code}] Après fusion vision+texte : {len(result.get('equipe', []))} membres")
        for key in ["adherents", "contacts"]:
            if not result.get(key) and vision_result.get(key):
                result[key] = vision_result[key]

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
