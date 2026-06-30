# 🎬 Mediathek Film Tracker — Arte × Letterboxd

**An automated weekly digest of feature films streaming on Germany's public-broadcast media libraries, ranked by their Letterboxd community rating.**

Every Friday a GitHub Actions workflow pulls the newest films from the public broadcasters (Arte, ZDF, ARD, 3sat), enriches each one with its Letterboxd rating, director and synopsis, and produces a colour-coded Excel report — so the best films currently available are visible at a glance, with a separate sheet highlighting new titles from a personal list of favourite directors.

---

## Why this project

Public media libraries host hundreds of films, but their interfaces make it hard to answer a simple question: *what's actually worth watching right now?* This tool answers it by combining three data sources into one ranked, filterable table and running entirely on its own schedule.

It demonstrates a small but complete data pipeline:

- **REST API consumption** with a non-obvious content-type requirement
- **Multi-source enrichment** — cross-referencing a German title against TMDb to reach the right Letterboxd page
- **HTML / JSON-LD scraping** of server-rendered metadata
- **Deduplication logic** that picks the best language version of each film
- **Styled report generation** with `openpyxl` (conditional colour-coding, multiple sheets, frozen headers, autofilters)
- **CI/CD automation** with GitHub Actions, including state that persists across scheduled runs

---

## How it works

The pipeline runs in five steps:

**1. Load the film list** — Films are fetched from the [MediathekViewWeb](https://mediathekviewweb.de/) API, which aggregates all German public-broadcast libraries. Results are filtered per channel, restricted to a minimum runtime of 60 minutes (to exclude short films), and sorted newest-first.

**2. Pick the best version** — The same film is often offered in up to five variants (dubbed, original-with-subtitles, audio description, etc.). For each title the pipeline keeps exactly one: the original version with German subtitles where available, otherwise the German dub — and discards the rest.

**3. Enrich via TMDb → Letterboxd** — Letterboxd loads its search results with JavaScript, so direct scraping fails. Instead the German title is looked up on TMDb to obtain a stable film ID, and `letterboxd.com/tmdb/{id}/` redirects straight to the correct film page. That page's JSON-LD block yields the rating, director, year and synopsis without needing a browser.

**4. Query every film** — Each candidate runs through the full pipeline with rate-limiting between requests to stay polite to both APIs.

**5. Build the report** — Results are sorted by rating and written to a styled Excel workbook. A rating-based colour scale flags the strongest films, and a second sheet collects anything directed by a name on the favourites list. A small CSV "database" tracks which films have already been seen, so each run can distinguish genuinely new titles from repeats.

---

## Tech stack

| Purpose | Library |
| --- | --- |
| HTTP requests | `requests` |
| HTML parsing | `beautifulsoup4`, `lxml` |
| Data handling | `pandas` |
| Excel generation | `openpyxl` |
| Automation | GitHub Actions |

External services: **MediathekViewWeb API** (no key required) and the **TMDb API** (free key required).

---

## Project structure

```
filmdatenbank-mediatheken/
├── main.py                      # Full pipeline as a runnable script
├── config.py                    # Channels, favourite directors, rating thresholds
├── requirements.txt             # Python dependencies
├── notebooks/
│   └── arte_letterboxd.ipynb    # Exploratory notebook version with sanity checks
├── data/                        # Persisted state (tracked in the repo)
│   ├── filme_db.csv             # Cumulative database of all films seen
│   └── letzte_woche.csv         # Last run's titles, for new-film detection
└── .github/workflows/
    └── mediathek-tracker.yml    # Weekly scheduled run
```

> Note: the code, comments and variable names are written in German; this README is in English.

---



## Automation (GitHub Actions)

The workflow runs every Friday at 09:00 UTC and can also be triggered manually from the **Actions** tab via *Run workflow*.

```yaml
on:
  schedule:
    - cron: '0 9 * * 5'   # Fridays, 09:00 UTC
  workflow_dispatch:
```

On each run the workflow:

1. Installs dependencies and runs `main.py` (the TMDb key is supplied from repository secrets).
2. Uploads the generated Excel as a downloadable **artifact**, retained for 30 days.
3. Commits the updated `data/` CSV files back to the repo, so the "already seen" state survives between runs.

To download a report, open the relevant run under the **Actions** tab and grab the file from the **Artifacts** section at the bottom.

> The workflow needs `permissions: contents: write` so the scheduled run can commit its state back. Because each run pushes a commit, remember to `git pull` before pushing local changes.

---

## Example output

A sample report is included in this repository: **`arte_letterboxd_20260618.xlsx`**.

The main sheet lists each film with the following columns, sorted by rating and colour-coded:

| Column | Content |
| --- | --- |
| Titel | Film title (Letterboxd / original) |
| Sender | Broadcasting channel |
| Jahr | Release year |
| Regie | Director |
| LB ★ | Letterboxd rating |
| Stimmen | Number of ratings |
| Dauer (min) | Runtime in minutes |
| Arte-Datum | Date added to the library |
| Beschreibung | Synopsis |
| Arte-Link / Letterboxd-Link | Direct links to both pages |

Rating-based colour scale:

| Colour | Rating | Meaning |
| --- | --- | --- |
| 🟢 Medium green | ★ > 4.0 | Very good |
| 🟩 Light green | ★ 3.66 – 4.0 | Good |
| 🟡 Light yellow | ★ 3.4 – 3.65 | Okay |

A second sheet, **Lieblingsregisseure** ("favourite directors"), filters the same data down to films directed by anyone on the configured favourites list.
