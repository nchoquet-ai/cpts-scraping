# Projet CPTS — Collecte et vérification des données

Pipeline complet de collecte et vérification des données des ~852 CPTS françaises.

## Fichiers du projet

| Fichier | Description |
|---|---|
| `verify_and_complete_sites.py` | Vérifie et complète les URLs des sites CPTS |
| `verify_and_complete_sites_PROCESS.txt` | Documentation détaillée du script de vérification |
| `scrape_cpts.py` | Scrape les sites CPTS et extrait les données (équipe, projets…) |
| `scrape_cpts_PROCESS.txt` | Documentation détaillée du script de scraping |
| `requirements.txt` | Dépendances Python à installer |

## Installation (une seule fois)

```bash
pip install -r requirements.txt
playwright install chromium
```

## Clé API Anthropic (à chaque session CMD)

```cmd
set ANTHROPIC_API_KEY=sk-ant-...
```

Pour ne plus avoir à la retaper, ajoutez-la en variable d'environnement permanente :
Panneau de configuration → Système → Variables d'environnement → Nouvelle

---

## Script 1 : verify_and_complete_sites.py

Vérifie les URLs existantes, confirme la correspondance CPTS, et recherche les sites manquants.

### Commandes

```cmd
# Lancement complet sur 852 CPTS (~1h10)
python verify_and_complete_sites.py --input "Input/Liste des CPTS avec site.xlsx"

# Test sur 50 lignes
python verify_and_complete_sites.py --input "Input/Liste des CPTS avec site.xlsx" --limit 50

# Reprendre après interruption
python verify_and_complete_sites.py --input "Input/Liste des CPTS avec site.xlsx" --resume

# Phase 1 & 2 uniquement (vérification sans recherche)
python verify_and_complete_sites.py --input "Input/Liste des CPTS avec site.xlsx" --verify-only

# Phase 3 uniquement (recherche des manquants)
python verify_and_complete_sites.py --input "Input/Liste des CPTS avec site.xlsx" --search-only
```

### Fichiers produits

```
Output/
├── CPTS_sites_verified.xlsx   ← Résultat principal (Excel coloré)
├── CPTS_sites_verified.csv    ← Même données en CSV
├── checkpoint.csv             ← Sauvegarde intermédiaire (reprise)
└── run.log                    ← Journal de l'exécution
```

### Durée estimée (852 CPTS)

| Phase | Durée |
|---|---|
| Phase 1 — Vérification HTTP | ~3 min |
| Phase 2 — Correspondance Claude | ~35 min |
| Pause rate limit | ~2 min |
| Phase 3 — Recherche manquants | ~30 min |
| **Total** | **~1h10** |

### Coût API estimé : ~3 à 5 USD

---

## Script 2 : scrape_cpts.py

Scrape les sites web des CPTS et extrait les informations structurées via Claude.

### Commandes

```cmd
# Charger la liste Excel et scraper (phase 1)
python scrape_cpts.py --scrape --input "Input/Liste des CPTS avec site.xlsx"

# Test sur 9 CPTS intégrées
python scrape_cpts.py --scrape --test

# Extraire les infos via Claude (phase 2)
python scrape_cpts.py --extract

# Re-extraire avec un prompt modifié
python scrape_cpts.py --extract --rerun

# Exporter en CSV
python scrape_cpts.py --export

# Voir l'état de la base
python scrape_cpts.py --status
```

### Architecture SQLite

Les données sont stockées dans `Output/cpts.db` :

| Table | Contenu |
|---|---|
| `cpts` | Liste des CPTS (nom, URL, métadonnées) |
| `pages` | Texte brut scraped (conservé indéfiniment) |
| `extractions` | Données structurées extraites par Claude |

**Avantage clé** : le scraping (lent) et l'extraction (rapide) sont séparés.
Pour ajouter un nouveau champ, modifiez `EXTRACTION_PROMPT` et relancez `--extract --rerun`
sans re-scraper les sites.

### Données extraites par CPTS

- Équipe dirigeante (président, directeur, vice-président…)
- Projets et missions en cours
- Coordonnées (téléphone, email, adresse)
- Communes couvertes
- Nombre de professionnels membres
- Date de création

---

## Prochaines étapes prévues

- [ ] Ajouter la collecte des **actualités** (Niveau 1 : pages actu des sites + Niveau 2 : recherche Claude)
- [ ] Ajouter la collecte des **partenariats** (MSP, centres de santé…)
- [ ] Croisement avec la base **RPPS** (code RPPS, spécialité, lieu d'exercice)
- [ ] Automatisation du run 2-3 fois par an

---

## Conformité RGPD

Les données collectées sont des **données professionnelles publiques** publiées
volontairement sur les sites officiels des CPTS (noms de dirigeants dans
l'exercice de leurs fonctions, coordonnées professionnelles, projets institutionnels).

Base légale : intérêt légitime (RGPD Art. 6.1.f).

Anthropic (sous-traitant) dispose d'un DPA conforme au RGPD avec Clauses
Contractuelles Types UE, automatiquement inclus dans les CGU commerciales.
