"""
Scraping CPTS — PHASE 1 : Capture brute exhaustive
===================================================
Visite TOUS les liens internes depth=1 de chaque site CPTS et capture
le contenu brut (DOM text + screenshot systématique) sans aucune analyse.

Principe :
  1. Charge la home page
  2. Extrait tous les liens internes (depth=1, sans limite, sans filtre)
  3. Visite chaque lien
  4. Pour chaque page : DOM text + screenshot full-page → stocké en DB
  5. Une ligne par page dans la table `pages`

AUCUNE analyse : pas d'appel Claude, pas de scoring, pas de tag,
pas d'extraction email, pas de détection organigramme.
L'analyse est faite en Phase 2 (script séparé).

Usage :
  pip install playwright pandas openpyxl tqdm
  playwright install chromium

  # Scraping complet
  python scrape_phase1.py --input "Input\\Liste des CPTS avec site validé.xlsx"

  # Test sur N CPTS
  python scrape_phase1.py --input "..." --limit 5

  # Reprise après interruption (sans reset des erreurs)
  python scrape_phase1.py --input "..." --resume

  # Stats DB
  python scrape_phase1.py --stats
"""

import argparse
import asyncio
import base64
import logging
import re
import sqlite3
import sys
from collections import deque
from pathlib import Path
from urllib.parse import urlparse

# Force UTF-8 output on Windows terminals (cp1252 can't encode emoji/box chars)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import httpx
import pandas as pd
from tqdm import tqdm

# ─── Configuration ─────────────────────────────────────────────────────────────

OUTPUT_DIR      = Path("Output")
DB_PATH         = OUTPUT_DIR / "cpts.db"
SCREENSHOTS_DIR = OUTPUT_DIR / "screenshots"

# Playwright
PAGE_TIMEOUT = 20_000   # ms
SCROLL_PAUSE = 2.0      # secondes après scroll (lazy-loading WordPress)
NAV_PAUSE    = 2.0      # secondes après navigation

# Parallélisation : nombre de CPTS scrapées simultanément (chacune = 1 contexte Chromium)
CONCURRENCY  = 8        # surchargé par --concurrency

# Découverte des pages : sitemap.xml + liens home + profondeur 2 (BFS), plafonné par CPTS
MAX_PAGES = 200         # plafond de pages visitées par CPTS (surchargé par --max-pages)
MAX_DEPTH = 2           # profondeur de crawl (0=home, 1=liens home/sitemap, 2=sous-liens)
ORG_RE    = re.compile(r"organigramme|trombinoscope|org.?chart", re.IGNORECASE)
SKIP_EXT  = re.compile(r"\.(pdf|jpg|jpeg|png|gif|svg|webp|css|js|zip|docx?|xlsx?|pptx?)$", re.IGNORECASE)
HTTP_UA   = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}

# ─── Logging ───────────────────────────────────────────────────────────────────

def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(OUTPUT_DIR / "scrape_phase1.log", mode="a", encoding="utf-8"),
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
        scrape_ts       TEXT
    );

    CREATE TABLE IF NOT EXISTS pages (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        code            TEXT,
        page_type       TEXT,
        url             TEXT,
        dom_text        TEXT,
        screenshot_path TEXT,
        used_vision     INTEGER DEFAULT 0,
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


def save_page(con, code, page_type, url, dom_text, screenshot_path=None, used_vision=False):
    cur = con.cursor()
    cur.execute(
        """INSERT INTO pages
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


def fetch_sitemap_urls(base_url: str, limit: int = 1000) -> list[str]:
    """
    Récupère la liste des pages depuis le sitemap.xml (et sous-sitemaps d'index).
    Retourne les URLs internes (même domaine, hors fichiers). Best-effort, jamais bloquant.
    """
    base_domain = urlparse(base_url).netloc
    root = f"{urlparse(base_url).scheme}://{base_domain}"
    found: list[str] = []
    seen_sm: set[str] = set()

    def _get(u):
        try:
            r = httpx.get(u, headers=HTTP_UA, timeout=15, follow_redirects=True)
            return r.text if r.status_code == 200 else ""
        except Exception:
            return ""

    candidates = deque([root + p for p in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml")])
    while candidates and len(found) < limit:
        sm = candidates.popleft()
        if sm in seen_sm:
            continue
        seen_sm.add(sm)
        xml = _get(sm)
        if not xml:
            continue
        locs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml, re.IGNORECASE)
        for loc in locs:
            if loc.lower().endswith(".xml"):           # sous-sitemap d'index
                if loc not in seen_sm and base_domain in loc:
                    candidates.append(loc)
            elif base_domain in loc and not SKIP_EXT.search(loc) and "#" not in loc:
                found.append(loc)
    # dédup en préservant l'ordre
    seen, out = set(), []
    for u in found:
        k = u.rstrip("/")
        if k not in seen:
            seen.add(k); out.append(u)
    return out

# ─── Phase 1 : Scraping Playwright ─────────────────────────────────────────────

async def _fetch_org_image(page) -> str | None:
    """
    Si la page contient un organigramme/trombinoscope (image), télécharge l'image
    NATIVE (haute résolution) via httpx et retourne son base64. Sinon None.
    Gère les images Wix servies en AVIF malgré l'extension PNG.
    """
    try:
        page_html = await page.content()
        if not ORG_RE.search(page_html):
            return None
        org_imgs = await page.evaluate("""() =>
            Array.from(document.querySelectorAll('img'))
                .filter(i => (i.alt && /organigramme|trombinoscope|org.?chart/i.test(i.alt))
                          || (i.src && /organigramme|trombinoscope|org.?chart/i.test(i.src)))
                .map(i => i.src)
        """)
        if not org_imgs:
            return None
        img_url = org_imgs[0]
        if "wixstatic.com" in img_url:
            img_url = re.sub(r"enc_avif[^/]*", "enc_png", img_url)
            img_url = re.sub(r"quality_auto[^/]*", "", img_url)
        r = httpx.get(img_url, headers=HTTP_UA, timeout=15, follow_redirects=True)
        if r.status_code != 200:
            return None
        if "avif" in r.headers.get("content-type", ""):
            base = img_url.split("/v1/")[0] + "/v1/fill/w_1200/" + img_url.split("/")[-1]
            r2 = httpx.get(base, headers=HTTP_UA, timeout=15, follow_redirects=True)
            if r2.status_code == 200:
                r = r2
        log.info(f"Image organigramme native récupérée ({r.headers.get('content-type','?')})")
        return base64.b64encode(r.content).decode()
    except Exception as e:
        log.debug(f"Erreur récupération organigramme : {e}")
        return None


async def get_page_content(page, url: str, skip_navigation: bool = False) -> tuple[str, str | None, str | None]:
    """
    Charge une URL et retourne (dom_text, screenshot_b64, org_image_b64_or_None).
    Screenshot SYSTÉMATIQUE. org_image = image organigramme native si détectée.
    Si skip_navigation=True, utilise le DOM déjà chargé (évite double goto).
    """
    try:
        if not skip_navigation:
            try:
                await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
            except Exception:
                try:
                    await page.goto(url, timeout=30000, wait_until="commit")
                except Exception as e:
                    log.warning(f"Erreur goto {url}: {e}")
                    return "", None, None
            await asyncio.sleep(NAV_PAUSE + 1)  # +1s pour Wix/WP

        # Scroll pour déclencher le lazy-loading
        try:
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await asyncio.sleep(SCROLL_PAUSE)
            await page.evaluate("window.scrollTo(0, 0)")
            await asyncio.sleep(2.0)
        except Exception:
            pass

        # Extraction texte DOM brut (script/style/nav/footer retirés)
        dom_text = await page.evaluate("""() => {
            const remove = ['script','style','nav','footer','header','.cookie','#cookie'];
            remove.forEach(sel => {
                try { document.querySelectorAll(sel).forEach(el => el.remove()); } catch(e) {}
            });
            return document.body.innerText.replace(/\\s+/g, ' ').trim();
        }""")

        # Image organigramme native (avant suppression DOM ? on relit le HTML complet)
        org_b64 = await _fetch_org_image(page)

        # Screenshot systématique full-page
        screenshot_bytes = await page.screenshot(full_page=True)
        screenshot_b64   = base64.b64encode(screenshot_bytes).decode()

        return dom_text, screenshot_b64, org_b64

    except Exception as e:
        log.warning(f"Erreur get_page_content {url}: {e}")
        return "", None, None


async def scrape_one(browser, code: str, nom: str, url: str, con) -> bool:
    """Scrape un site CPTS complet : home + tous les liens internes depth=1."""
    context = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="fr-FR",
    )
    page = await context.new_page()

    base_domain = urlparse(url).netloc

    def internal(h: str) -> bool:
        return (base_domain in h) and ("#" not in h) and not SKIP_EXT.search(h)

    async def harvest_links(p) -> list[str]:
        try:
            return await p.evaluate(f"""() => {{
                const domain = '{base_domain}';
                return [...new Set(
                    Array.from(document.querySelectorAll('a[href]'))
                        .map(a => a.href)
                        .filter(h => h.includes(domain) && !h.includes('#'))
                )];
            }}""")
        except Exception:
            return []

    try:
        log.info(f"[{code}] Scraping {url}")
        await page.goto(url, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
        await asyncio.sleep(NAV_PAUSE)
        try:
            await page.wait_for_selector("body", timeout=5000)
        except Exception:
            pass

        # ── Seeds : home + liens home (depth 1) ∪ sitemap (depth 1) ───────────
        home_links = await harvest_links(page)
        sitemap_links = fetch_sitemap_urls(url)
        log.info(f"[{code}] {len(home_links)} liens home | {len(sitemap_links)} pages sitemap")

        seen: set[str] = set()
        queue: deque = deque()

        def enqueue(u: str, depth: int):
            k = u.rstrip("/")
            if k and k not in seen and internal(u):
                seen.add(k); queue.append((u, depth))

        enqueue(url, 0)
        for l in home_links:    enqueue(l, 1)
        for l in sitemap_links: enqueue(l, 1)

        # ── BFS jusqu'à MAX_DEPTH, plafonné à MAX_PAGES ───────────────────────
        visited = 0
        idx = 0
        while queue and visited < MAX_PAGES:
            page_url, depth = queue.popleft()
            skip = (page_url.rstrip("/") == url.rstrip("/") and visited == 0)
            dom_text, ss_b64, org_b64 = await get_page_content(page, page_url, skip_navigation=skip)

            ss_dir = SCREENSHOTS_DIR / code
            ss_dir.mkdir(parents=True, exist_ok=True)

            ss_path = None
            if ss_b64:
                ss_path = str(ss_dir / f"{code}_{idx}.png")
                with open(ss_path, "wb") as f:
                    f.write(base64.b64decode(ss_b64))
            save_page(con, code, "", page_url, dom_text,
                      screenshot_path=ss_path, used_vision=bool(ss_b64))

            # Image organigramme native → ligne dédiée page_type='organigramme'
            if org_b64:
                org_path = str(ss_dir / f"{code}_{idx}_org.png")
                with open(org_path, "wb") as f:
                    f.write(base64.b64decode(org_b64))
                save_page(con, code, "organigramme", page_url, "",
                          screenshot_path=org_path, used_vision=True)

            visited += 1
            idx += 1

            # Profondeur : récupérer les sous-liens (depth+1) tant que < MAX_DEPTH
            if depth < MAX_DEPTH:
                for l in await harvest_links(page):
                    enqueue(l, depth + 1)

        log.info(f"[{code}] {visited} pages scrapées (plafond {MAX_PAGES})")
        set_scrape_status(con, code, "done")
        log.info(f"[{code}] ✅ Scraping OK")
        return True

    except Exception as e:
        log.error(f"[{code}] ❌ Erreur scraping: {e}")
        set_scrape_status(con, code, "error")
        return False
    finally:
        await context.close()


async def run_scraping(con, limit=None, resume=False, concurrency=CONCURRENCY):
    """
    Lance le scraping Playwright en parallèle sur les CPTS pending.
    `concurrency` CPTS sont scrapées simultanément (1 contexte Chromium chacune).
    """
    from playwright.async_api import async_playwright

    if not resume:
        # Remettre les erreurs en pending pour retry
        con.execute("UPDATE cpts SET scrape_status='pending' WHERE scrape_status='error'")
        con.commit()

    todo = get_pending_scrape(con, limit)
    log.info(f"Phase 1 — {len(todo)} CPTS à scraper (concurrence={concurrency})")

    if not todo:
        log.info("Rien à scraper.")
        return

    sem      = asyncio.Semaphore(concurrency)
    progress = tqdm(total=len(todo), desc="Scraping")
    counter  = {"done": 0}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)

        async def worker(code, nom, url):
            async with sem:
                await scrape_one(browser, code, nom, url, con)
                progress.update(1)
                counter["done"] += 1
                # Checkpoint tous les 25 sites terminés
                if counter["done"] % 25 == 0:
                    ok  = con.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='done'").fetchone()[0]
                    err = con.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='error'").fetchone()[0]
                    log.info(f"Checkpoint {counter['done']}/{len(todo)} — ✅ {ok} scrapées | ❌ {err} erreurs")

        await asyncio.gather(*(worker(code, nom, url) for code, nom, url in todo))
        await browser.close()

    progress.close()
    log.info("Phase 1 terminée.")

# ─── Statistiques ───────────────────────────────────────────────────────────────

def print_stats(con):
    cur = con.cursor()
    total   = cur.execute("SELECT COUNT(*) FROM cpts").fetchone()[0]
    scraped = cur.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='done'").fetchone()[0]
    errors  = cur.execute("SELECT COUNT(*) FROM cpts WHERE scrape_status='error'").fetchone()[0]
    pages   = cur.execute("SELECT COUNT(*) FROM pages").fetchone()[0]

    print("\n─── Résumé Phase 1 ─────────────────────────────────────")
    print(f"  Total CPTS          : {total}")
    print(f"  ✅ Scrapées         : {scraped}")
    print(f"  ❌ Erreurs scraping : {errors}")
    print(f"  📄 Pages capturées  : {pages}")
    print("────────────────────────────────────────────────────────\n")

# ─── Pipeline principal ─────────────────────────────────────────────────────────

async def run(args):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    con = init_db()

    if not args.input:
        log.error("--input requis pour le scraping.")
        return

    df = pd.read_excel(args.input, dtype=str).fillna("")
    df.columns = [c.strip() for c in df.columns]

    url_col  = next((c for c in df.columns if any(k in c.lower() for k in ["site", "url", "lien"])), None)
    nom_col  = next((c for c in df.columns if any(k in c.lower() for k in ["nom", "label", "libellé"])), None)
    code_col = next((c for c in df.columns if any(k in c.lower() for k in ["code", "id"])), df.columns[0])

    if not url_col or not nom_col:
        log.error(f"Colonnes non trouvées. Disponibles : {df.columns.tolist()}")
        return

    log.info(f"Colonnes détectées → code: '{code_col}' | nom: '{nom_col}' | url: '{url_col}'")

    rows = []
    for _, row in df.iterrows():
        url = normalize_url(row.get(url_col, ""))
        if url:
            rows.append((row[code_col], row[nom_col], url))

    upsert_cpts(con, rows)
    log.info(f"{len(rows)} CPTS chargées (avec URL valide)")

    if args.limit:
        all_codes = [r[0] for r in rows]
        keep = set(all_codes[:args.limit])
        con.execute(
            f"UPDATE cpts SET scrape_status='skip' WHERE code NOT IN ({','.join('?'*len(keep))})",
            list(keep)
        )
        con.commit()

    await run_scraping(con, limit=args.limit, resume=args.resume,
                       concurrency=args.concurrency)

    print_stats(con)
    con.close()

# ─── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Scraping CPTS — Phase 1 (capture brute)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input",  help="Fichier Excel source")
    parser.add_argument("--limit",  type=int, help="Limiter aux N premières CPTS")
    parser.add_argument("--resume", action="store_true", help="Reprendre sans reset des erreurs")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY,
                        help=f"CPTS scrapées en parallèle (défaut: {CONCURRENCY})")
    parser.add_argument("--max-pages", type=int, default=MAX_PAGES, dest="max_pages",
                        help=f"Plafond de pages par CPTS (défaut: {MAX_PAGES})")
    parser.add_argument("--stats",  action="store_true", help="Afficher les stats de la DB")
    args = parser.parse_args()

    MAX_PAGES = args.max_pages

    if not any([args.input, args.stats]):
        parser.print_help()
    else:
        setup_logging()
        if args.stats and not args.input:
            con = init_db()
            print_stats(con)
            con.close()
        else:
            asyncio.run(run(args))
