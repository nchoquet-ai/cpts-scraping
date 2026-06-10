# État du projet CPTS Scraping — handoff

_Dernière mise à jour : 10 juin 2026._

Pipeline en 2 phases pour collecter et structurer les données de ~852 CPTS françaises.
Lot pilote courant = **12 CPTS** (fichier `Input/Liste des CPTS avec site validé.xlsx`).

## Architecture

```
scrape_phase1.py  → capture brute (Playwright) : sitemap + liens home + profondeur 2,
                    screenshot systématique + DOM, image organigramme native (httpx).
                    Écrit DB Output/cpts.db (tables cpts, pages). AUCUN appel Claude.
scrape_phase2.py  → analyse Claude (holistique, sans routage) : équipe + projets/missions
                    + actus + adhérents + contacts, avec source_url + dates.
                    Mapping multi-axes + codage CIM-10 cascade. Export 4 CSV.
```

DB unique `Output/cpts.db`. `Input/` et `Output/` sont **gitignorés** (données locales).

## Référentiel de catégories (`Input/CPTS_categories_config.xlsx`, 5 onglets)
- **A** Mission/parcours (aligné ACI) — 8
- **B** Pathologies (CIM-10 niveau 3) — 25 curées + cascade (curé → fichier officiel 2054 → hors-réf)
- **C** Public/cycle de vie — 6
- **D** Vie de la CPTS — 6
- **E** Situation/vulnérabilité — 5

Fichier CIM-10 officiel : `Input/CIM10_FR2025_niveau3.csv` (2054 codes, parsé du ClaML ATIH 2025).
Plan détaillé : `C:\Users\nicol\.claude\plans\concurrent-bouncing-pearl.md`.

## Mécanisme Axe B (pathologies)
Claude détecte la pathologie en **texte libre** → codage cascade déterministe :
table curée → rapprochement fuzzy sur les 2054 libellés officiels (flag « à vérifier ») → hors-réf.
Rapport d'enrichissement : `Output/pathologies_a_enrichir.csv`. Boucle assistée : on ajoute
les pathos récurrentes à l'onglet B, puis re-mapping déterministe (sans re-appeler Claude).

## Schéma d'extraction (champs)
- Équipe : civilité, prénom, nom, fonction, spécialité, source_url
- Adhérents : nombre (le plus récent) + source_url
- Contacts : email, téléphone, adresse, source_url
- Projets/Missions (fusionnés) : titre, description, date, pathologies[], source_url + axes A/B/C/D/E
- Actus : titre, date, résumé, pathologies[], source_url + axes

## État d'avancement
- ✅ Phase 1 (brute) + parallélisation (`--concurrency`, défaut 8) + checkpoint /25
- ✅ Phase 1 v2 : découverte sitemap + profondeur 2 (`--max-pages`, défaut 200) + organigramme natif
- ✅ Phase 2 : extraction holistique + Vision ciblée (mots-clés/URL équipe + DOM pauvre + organigramme prioritaire),
  redimensionnement screenshots ≤7500px, filtre pages-coquilles (<200 chars), codage CIM-10, export CSV
- ✅ Référentiel + fichier CIM-10 officiel
- ✅ Validé sur lot pilote :
  - Phase 1 v2 ARA-5=112 pages (organigramme), HDF-18=200 pages (/lequipe/ capté)
  - Phase 2 : ARA-5 17 membres (organigramme natif), HDF-18 58 membres (/lequipe/),
    NA-73 ~25, PDL-7 ~13, PDL-24 ~55 actus, etc.

## Prochaines étapes (tâches ouvertes)
1. **Tuning prompt équipe** — HDF-18=58 est sur-extrait (inclut l'annuaire des professionnels).
   Contraindre EXTRACT_PROMPT à ne garder que le bureau/CA/gouvernance. (tâche #10)
2. **Mode `--batch`** (Batches API, −50 % coût) à ajouter à scrape_phase2 **avant le run prod**. (tâche #9)
3. **Run production 850 CPTS** quand validé. Estimation coût Phase 2 : ~80-125 $ en batch.

## Commandes
```bash
# Phase 1
python scrape_phase1.py --input "Input\Liste des CPTS avec site validé.xlsx" --concurrency 8 --max-pages 200
# Phase 2 (clé API requise : export ANTHROPIC_API_KEY=...)
python scrape_phase2.py --extract --export            # tout
python scrape_phase2.py --extract --export --limit 11 # test pilote
python scrape_phase2.py --stats
```

## Cas particuliers connus
- PACA-20 : domaine DNS mort (cptspaysdazur.fr) → reste en `error`.
- Sites image-only (organigramme) : nécessitent le téléchargement natif + Vision.
- Pages orphelines (non liées au menu) : captées via sitemap.
- Variabilité LLM : les comptes membres/projets fluctuent légèrement entre runs.
