#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
sys.path.insert(0, '.')
from config import SENDER, LIEBLINGSREGISSEURE

import requests
import json
import re
import time
import urllib.parse
from datetime import datetime
from collections import defaultdict

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Farben für die drei Bewertungsstufen ───────────────────────────────
FARBE_GUT       = "D6EFD8"   # Hellgrün  — 3.66 bis 4.0
FARBE_SEHR_GUT  = "A8D5A2"   # Mittelgrün — über 4.0
FARBE_OK        = "FFF3CD"   # Hellgelb  — 3.4 bis 3.65

# ── Konfiguration ──────────────────────────────────────────────────────────
# Maximale Anzahl Filme für die Letterboxd-Abfrage (für Tests klein halten)
MAX_FILMS = None

# Pause zwischen Letterboxd-Anfragen — bitte nicht unter 0.5 setzen
LETTERBOXD_DELAY = 0.8

# Mindestlaufzeit in Sekunden (3600 = 60 Minuten → schließt Kurzfilme aus)
MIN_DURATION_SEC = 3600

# TMDb API-Key (kostenlos von themoviedb.org/settings/api)
TMDB_KEY = os.environ.get('TMDB_KEY', '')
if not TMDB_KEY:
    raise ValueError("TMDB_KEY Umgebungsvariable nicht gesetzt!")

print("✓ Imports und Konfiguration geladen")

# ---
# Schritt 1: Arte-Filmliste von MediathekViewWeb laden
#
# MediathekViewWeb aggregiert alle öffentlich-rechtlichen Mediatheken und bietet eine Such-API.
#
# Wir filtern nach:
# - Sender: ARTE.DE
# - Thema: enthält "film" → trifft Kino - Filme, Fernsehfilm, etc.
# - Mindestdauer: 3600 Sekunden (60 Min.) → schließt Kurzfilme aus
#
# ⚠️ Wichtig: Content-Type muss text/plain sein, sonst antwortet die API mit 400.
#

def lade_mediathek_filme(sender, topic="film", size=500, duration_min=MIN_DURATION_SEC):
    """
    Lädt Filme von MediathekViewWeb.
    Gibt eine Liste von Einträgen zurück (jeder Eintrag = eine Versionsvariant eines Films).
    """
    query = {
        "queries": [
            {"fields": ["channel"], "query": sender},
            {"fields": ["topic"],   "query": topic}
        ],
        "sortBy":       "timestamp",
        "sortOrder":    "desc",
        "future":       False,
        "offset":       0,
        "size":         size,
        "duration_min": duration_min
    }

    antwort = requests.post(
        "https://mediathekviewweb.de/api/query",
        data=json.dumps(query),
        headers={"Content-Type": "text/plain"},  # MUSS text/plain sein
        timeout=30
    )
    antwort.raise_for_status()

    data = antwort.json()
    if data.get("err"):
        raise ValueError(f"API-Fehler: {data['err']}")

    result_info = data.get("result", {})
    total  = result_info.get("queryInfo", {}).get("totalResults", "?")
    filme  = result_info.get("results", [])

    print(f"Treffer gesamt in API : {total}")
    print(f"Zurückgegeben         : {len(filme)}")
    return filme


rohdaten = []
for sender, topic in SENDER.items():
    print(f'Lade {sender}...')
    treffer = lade_mediathek_filme(sender, topic)
    print(f'  → {len(treffer)} Treffer')
    rohdaten.extend(treffer)
print(f'Gesamt: {len(rohdaten)} Einträge')

# Sanity Check: Treffer pro Sender
df_check = pd.DataFrame(rohdaten)
print('── Treffer pro Sender ──────────────')
print(df_check.groupby('channel')['title'].count())
print(f'\nGesamt: {len(rohdaten)} Einträge')
print(f'Eindeutig nach Filter: wird nach waehle_beste_version() angezeigt')

# ✅ Sanity Check 1: Hat die API sinnvolle Daten geliefert?
# Prüfe ob Daten vorhanden und vollständig sind
assert len(rohdaten) > 0, "API hat 0 Ergebnisse zurückgegeben — Verbindung prüfen"

# Zeige alle vorhandenen Suffixe (wichtig für den nächsten Schritt)
alle_suffixe = set()
for film in rohdaten:
    m = re.search(r'\(([^)]+)\)\s*$', film['title'])
    alle_suffixe.add(m.group(1) if m else "[kein Suffix]")

print(f"✓ {len(rohdaten)} Einträge geladen")
print(f"\nVorhandene Titel-Suffixe: {sorted(alle_suffixe)}")
print(f"\nBeispiel-Einträge (erste 6 Titel):")
for f in rohdaten[:6]:
    print(f"  {f['title']}")

# ---
# Schritt 2: Beste Version pro Film auswählen
#
# Arte stellt denselben Film in bis zu vier Varianten bereit — alle als separate Einträge in der API:
#
# | Suffix | Bedeutung | Wollen wir? |
# |---|---|---|
# | (Originalversion mit Untertitel) | Originalsprache + deutsche UT | ✅ 1. Wahl |
# | (kein Suffix) | Deutsche Synchronfassung | ✅ Fallback |
# | (mit Untertitel) | Deutsche Synchro + Untertitel für Gehörlose | ❌ |
# | (Audiodeskription) | Deutsche Synchro + Audiobeschreibung | ❌ |
# | (Originalversion) | Originalsprache ohne Untertitel | ❌ |
#
# Strategie: Für jeden Basis-Titel wählen wir genau eine Version aus — Priorität oben nach unten.
#
# ⚠️ Geändert gegenüber ursprünglichem Notebook: dedupliziere() und ist_omu() werden durch
# waehle_beste_version() ersetzt, das beides in einem Schritt erledigt.
#

def extrahiere_basis(titel):
    """
    Entfernt den Versions-Suffix in Klammern am Titelende.
    'Minari (Originalversion mit Untertitel)' → 'Minari'
    """
    return re.sub(
        r'\s*\((Originalversion mit Untertitel|Originalversion|mit Untertitel|Audiodeskription)\)\s*$',
        '',
        titel,
        flags=re.IGNORECASE
    ).strip()


def versionstyp(titel):
    """
    Klassifiziert einen Titel nach Versionstyp.

    Rückgabewerte:
        'original'  → (Originalversion mit Untertitel)  ← bevorzugt
        'deutsch'   → kein Suffix                       ← Fallback
        'skip'      → alles andere                      ← wird ignoriert
    """
    if re.search(r'\(Originalversion mit Untertitel\)', titel, re.IGNORECASE):
        return 'original'
    elif re.search(r'\(mit Untertitel\)|\(Audiodeskription\)|\(Originalversion\)', titel, re.IGNORECASE):
        return 'skip'
    else:
        return 'deutsch'   # kein Suffix = deutsche Synchronfassung


def bereinige_titel(titel):
    """
    Entfernt Suffix für externe Suchanfragen (TMDb, Letterboxd).
    Identisch mit extrahiere_basis(), als eigene Funktion für Klarheit.
    """
    return extrahiere_basis(titel)


def waehle_beste_version(filme):
    """
    Ersetzt dedupliziere() + ist_omu() aus dem ursprünglichen Notebook.

    Für jeden Basis-Titel pro Sender:
      1. (Originalversion mit Untertitel) wenn vorhanden
      2. sonst: kein Suffix (deutsche Fassung)
      3. alle anderen Varianten werden verworfen

    Bei mehreren Einträgen desselben Typs (Mehrfachausstrahlung)
    wird der neueste (höchster timestamp) behalten.

    Gleicher Film bei verschiedenen Sendern wird als separate Einträge behalten.
    """
    gruppen = defaultdict(dict)  # {(basis_titel, sender): {'original': film, 'deutsch': film}}

    for film in filme:
        basis = extrahiere_basis(film['title'])
        sender = film.get('channel', '')
        typ   = versionstyp(film['title'])

        if typ == 'skip':
            continue

        # Neuesten Eintrag bevorzugen
        vorhandener = gruppen[(basis, sender)].get(typ)
        if vorhandener is None or film['timestamp'] > vorhandener['timestamp']:
            gruppen[(basis, sender)][typ] = film

    ergebnis = []
    stats = {'original': 0, 'deutsch': 0}

    for basis_sender, versionen in gruppen.items():
        if 'original' in versionen:
            ergebnis.append(versionen['original'])
            stats['original'] += 1
        elif 'deutsch' in versionen:
            ergebnis.append(versionen['deutsch'])
            stats['deutsch'] += 1

    print(f"Rohdaten gesamt              : {len(filme)}")
    print(f"Eindeutige Basis-Titel+Sender: {len(gruppen)}")
    print(f"  davon Originalversion+UT   : {stats['original']}")
    print(f"  davon deutsche Fassung     : {stats['deutsch']}")
    print(f"  davon verworfen (skip)     : {len(filme) - sum(stats.values())}")
    return ergebnis


gefilterte_filme = waehle_beste_version(rohdaten)

# ✅ Sanity Check 2: Wurde die richtige Version pro Film gewählt?
assert len(gefilterte_filme) > 0, "Keine Filme nach Versionsfilter — waehle_beste_version() prüfen"

# Überprüfe: Gibt es noch unerwünschte Versionen?
unerwuenscht = [f for f in gefilterte_filme
                if re.search(r'\(mit Untertitel\)|\(Audiodeskription\)|\(Originalversion\)(?! mit)',
                             f['title'], re.IGNORECASE)]
assert len(unerwuenscht) == 0, f"Unerwünschte Versionen noch vorhanden: {[f['title'] for f in unerwuenscht]}"

# Zeige Beispiele beider Typen
original_bsp = [f for f in gefilterte_filme if '(Originalversion mit Untertitel)' in f['title']][:4]
deutsch_bsp  = [f for f in gefilterte_filme if '(Originalversion mit Untertitel)' not in f['title']][:4]

print(f"✓ {len(gefilterte_filme)} Filme nach Versionsfilter — alle Versionen korrekt\n")
print("Beispiele — Originalversion mit Untertitel:")
for f in original_bsp:
    print(f"  ✓ {f['title']}")
print("\nBeispiele — Deutsche Fassung (kein OmU verfügbar):")
for f in deutsch_bsp:
    print(f"  · {f['title']}")

# ---
# Schritt 3: Letterboxd-Daten via TMDb holen
#
# Letterboxd lädt seine Suchergebnisse per JavaScript — einfaches Scraping der Suchseite funktioniert nicht.
#
# Lösung in zwei Schritten:
#
# 1. TMDb API → sucht nach deutschem Titel → liefert tmdb_id + Originaltitel + Beschreibung
# 2. Letterboxd TMDb-Redirect → letterboxd.com/tmdb/{id}/ leitet direkt zur Filmseite weiter
#    → Filmseite enthält JSON-LD mit Bewertung, Regie, Jahr (server-seitig gerendert, kein JS)
#
# ⚠️ Geändert gegenüber ursprünglichem Notebook: suche_lb_slug() und hole_lb_filmseite()
# werden durch einen neuen dreistufigen Ablauf ersetzt. Letterboxd-Suche entfällt komplett.
#

# Browser-ähnliche Session — reduziert Bot-Erkennung bei Letterboxd
session = requests.Session()
session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
})


def hole_tmdb(deutscher_titel):
    """
    Sucht einen Film auf TMDb anhand des deutschen Titels.

    Rückgabe: (tmdb_id, original_title, jahr, deutsche_beschreibung)
    oder      (None, None, None, None) wenn nicht gefunden.

    TMDb findet "Der Mann aus Marmor" und gibt uns tmdb_id=9452 zurück,
    mit der wir direkt auf letterboxd.com/tmdb/9452/ zugreifen können.
    """
    r = requests.get(
        "https://api.themoviedb.org/3/search/movie",
        params={
            "api_key":  TMDB_KEY,
            "query":    deutscher_titel,
            "language": "de-DE",   # Suche auf Deutsch → bessere Trefferquote
        },
        timeout=10
    )
    r.raise_for_status()
    ergebnisse = r.json().get("results", [])

    if not ergebnisse:
        return None, None, None, None

    treffer = ergebnisse[0]
    return (
        treffer["id"],
        treffer.get("original_title", deutscher_titel),
        str(treffer.get("release_date", ""))[:4],
        treffer.get("overview", ""),
    )


def hole_lb_filmseite_von_html(html, url):
    soup = BeautifulSoup(html, "lxml")
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or ""
        raw = re.sub(r'/\*.*?\*/', '', raw, flags=re.DOTALL).strip()
        try:
            data = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            continue
        if data.get("@type") != "Movie":
            continue
        regie = ", ".join(
            d.get("name", "") for d in data.get("director", []) if d.get("name")
        )
        rat = data.get("aggregateRating", {})
        return {
            "lb_slug":     url.rstrip("/").split("/")[-1],
            "lb_title":    data.get("name", ""),
            "lb_year":     str(data.get("datePublished", ""))[:4],
            "lb_director": regie,
            "lb_desc":     data.get("description", ""),
            "lb_rating":   rat.get("ratingValue"),
            "lb_votes":    rat.get("ratingCount"),
            "lb_url":      url,
        }
    return None


def hole_lb_ueber_tmdb(tmdb_id):
    url = f"https://letterboxd.com/tmdb/{tmdb_id}/"
    try:
        r = session.get(url, timeout=15, allow_redirects=True)
        r.raise_for_status()
    except requests.RequestException:
        return None
    return hole_lb_filmseite_von_html(r.text, r.url)


def hole_lb_daten(film):
    titel = bereinige_titel(film.get("title", ""))
    time.sleep(0.3)
    tmdb_id, _, _, tmdb_desc = hole_tmdb(titel)
    if not tmdb_id:
        return {}
    time.sleep(LETTERBOXD_DELAY)
    lb = hole_lb_ueber_tmdb(tmdb_id)
    if lb and not lb.get("lb_desc"):
        lb["lb_desc"] = tmdb_desc
    return lb or {}


print("✓ Letterboxd-Funktionen bereit")

# ---
# Schritt 4: Alle Filme abfragen
#
# Jetzt wird für jeden Film in gefilterte_filme die komplette Pipeline durchlaufen.
#
# Laufzeit: ca. MAX_FILMS × 1.1 Sekunden (wegen Rate Limiting)
# Bei MAX_FILMS = 10 ca. 11 Sekunden.
#

# Sortiere nach Dauer (längste zuerst = wahrscheinlicher Spielfilm)
kandidaten = sorted(gefilterte_filme, key=lambda f: f.get("duration", 0), reverse=True)

if MAX_FILMS:
    kandidaten = kandidaten[:MAX_FILMS]

print(f"Starte Abfrage für {len(kandidaten)} Filme...")
print(f"Geschätzte Dauer: {len(kandidaten) * 1.1 / 60:.1f} Minuten\n")

ergebnisse = []

for i, film in enumerate(kandidaten, 1):
    titel = bereinige_titel(film.get("title", ""))
    dauer = f"{film.get('duration', 0) // 60} min"

    print(f"[{i:3d}/{len(kandidaten)}] {titel[:45]:45s} ({dauer})", end=" ")

    lb_daten = hole_lb_daten(film)

    if lb_daten:
        rating = lb_daten.get("lb_rating", "–")
        regie  = (lb_daten.get("lb_director") or "–")[:25]
        jahr   = lb_daten.get("lb_year", "–")
        print(f"→ ★ {rating}  {jahr}  {regie}")
    else:
        print("→ nicht gefunden")

    ergebnis = {
        "arte_titel":   film.get("title", ""),
        "arte_dauer":   film.get("duration", 0) // 60,
        "arte_datum":   datetime.fromtimestamp(film.get("timestamp", 0)).strftime("%d.%m.%Y"),
        "arte_thema":   film.get("topic", ""),
        "arte_url":     film.get("url_website", ""),
        "channel":      film.get("channel", ""),
        **lb_daten,
    }
    ergebnisse.append(ergebnis)

gefunden = sum(1 for e in ergebnisse if e.get("lb_rating"))
print(f"\n✓ Fertig: {gefunden} von {len(ergebnisse)} Filmen auf Letterboxd gefunden")

# ---
# Schritt 5: Ergebnisse aufbereiten und speichern
#

# Nicht gefundene Filme entfernen (kein lb_rating = nicht auf Letterboxd)
ergebnisse = [e for e in ergebnisse if e.get("lb_rating")]

print(f"✓ {len(ergebnisse)} Filme mit Letterboxd-Bewertung behalten")

# Lieblingsregisseure filtern
def ist_lieblingsregisseur(regie):
    if not regie:
        return False
    regie_lower = regie.lower()
    return any(name.lower() in regie_lower for name in LIEBLINGSREGISSEURE)


favoriten = [e for e in ergebnisse if ist_lieblingsregisseur(e.get("lb_director"))]

print(f"✓ {len(favoriten)} Filme von Lieblingsregisseuren gefunden:\n")
for f in sorted(favoriten, key=lambda e: e.get("lb_director", "")):
    print(f"  {f.get('lb_director'):25s}  ★ {f.get('lb_rating') or '–'}  {f.get('lb_title')}")

# Bewertungsfarbe bestimmen
def bewertungsfarbe(rating):
    if rating is None:
        return None
    if rating > 4.0:
        return FARBE_SEHR_GUT
    elif rating >= 3.66:
        return FARBE_GUT
    elif rating >= 3.4:
        return FARBE_OK
    return None

# ── Spalten die gespeichert werden ─────────────────────────────────────
SPALTEN = [
    ("lb_title",    "Titel"),
    ("channel",     "Sender"),
    ("lb_year",     "Jahr"),
    ("lb_director", "Regie"),
    ("lb_rating",   "LB ★"),
    ("lb_votes",    "Stimmen"),
    ("arte_dauer",  "Dauer (min)"),
    ("arte_datum",  "Arte-Datum"),
    ("lb_desc",     "Beschreibung"),
    ("arte_url",    "Arte-Link"),
    ("lb_url",      "Letterboxd-Link"),
]

# ── Workbook aufbauen ──────────────────────────────────────────────────
wb = Workbook()
ws = wb.active
ws.title = "Arte × Letterboxd"

thin = Side(style="thin", color="CCCCCC")
rahmen = Border(left=thin, right=thin, top=thin, bottom=thin)

# Kopfzeile
header_fill = PatternFill("solid", fgColor="1C1C2E")
for col, (_, anzeigename) in enumerate(SPALTEN, 1):
    zelle = ws.cell(row=1, column=col, value=anzeigename)
    zelle.font      = Font(name="Arial", bold=True, color="F5C842", size=10)
    zelle.fill      = header_fill
    zelle.alignment = Alignment(horizontal="center", vertical="center")
    zelle.border    = rahmen

ws.row_dimensions[1].height = 22

# Datenzeilen — nach Bewertung sortiert
sortiert = sorted(ergebnisse, key=lambda e: e.get("lb_rating") or 0, reverse=True)

for zeile_nr, eintrag in enumerate(sortiert, 2):
    rating = eintrag.get("lb_rating")
    farbe  = bewertungsfarbe(rating)
    fill   = PatternFill("solid", fgColor=farbe) if farbe else None

    for col, (key, _) in enumerate(SPALTEN, 1):
        wert = eintrag.get(key, "")

        # Zahlenformatierung
        if key == "lb_rating" and wert:
            wert = round(float(wert), 2)
        elif key == "lb_votes" and wert:
            wert = int(wert)

        zelle = ws.cell(row=zeile_nr, column=col, value=wert)
        zelle.font      = Font(name="Arial", size=9)
        zelle.alignment = Alignment(vertical="center", wrap_text=(key == "lb_desc"))
        zelle.border    = rahmen
        if fill:
            zelle.fill = fill

        # Bewertungsspalte fett
        if key == "lb_rating":
            zelle.font = Font(name="Arial", size=9, bold=True)
            zelle.number_format = "0.00"

        # Stimmen mit Tausenderpunkt
        if key == "lb_votes":
            zelle.number_format = "#,##0"

    # Zeilenhöhe an Beschreibungslänge anpassen
    zeichen = len(str(eintrag.get("lb_desc") or ""))
    zeilen  = max(1, -(-zeichen // 80))   # Aufrunden: wie viele Zeilen bei 80 Zeichen Breite
    ws.row_dimensions[zeile_nr].height = zeilen * 13 + 4

# ── Spaltenbreiten ──────────────────────────────────────────────────────
breiten = {
    "lb_title":    28,
    "channel":      8,
    "lb_year":      6,
    "lb_director": 22,
    "lb_rating":    8,
    "lb_votes":    12,
    "arte_dauer":   8,
    "arte_datum":  12,
    "lb_desc":     50,
    "arte_url":    14,
    "lb_url":      14,
}
for col, (key, _) in enumerate(SPALTEN, 1):
    ws.column_dimensions[get_column_letter(col)].width = breiten.get(key, 12)

# ── Autofilter + Fixierung ──────────────────────────────────────────────
ws.auto_filter.ref = f"A1:{get_column_letter(len(SPALTEN))}1"
ws.freeze_panes    = "A2"

# ── Legende ────────────────────────────────────────────────────────────
legende_zeile = len(sortiert) + 3
for farbe_hex, text in [
    (FARBE_SEHR_GUT, "★ > 4.0   — Sehr gut"),
    (FARBE_GUT,      "★ 3.66–4.0 — Gut"),
    (FARBE_OK,       "★ 3.4–3.65 — Okay"),
]:
    zelle = ws.cell(row=legende_zeile, column=1, value=text)
    zelle.fill = PatternFill("solid", fgColor=farbe_hex)
    zelle.font = Font(name="Arial", size=9)
    legende_zeile += 1

# ── Zweites Tabellenblatt: Lieblingsregisseure ─────────────────────────
ws2 = wb.create_sheet(title="Lieblingsregisseure")

# Kopfzeile (gleiche Spalten wie Haupttabelle)
for col, (_, anzeigename) in enumerate(SPALTEN, 1):
    zelle = ws2.cell(row=1, column=col, value=anzeigename)
    zelle.font      = Font(name="Arial", bold=True, color="F5C842", size=10)
    zelle.fill      = PatternFill("solid", fgColor="1C1C2E")
    zelle.alignment = Alignment(horizontal="center", vertical="center")
    zelle.border    = rahmen
ws2.row_dimensions[1].height = 22

# Daten — nach Regisseur dann Bewertung sortiert
favoriten_sortiert = sorted(
    favoriten,
    key=lambda e: (e.get("lb_director") or "", -(e.get("lb_rating") or 0))
)

for zeile_nr, eintrag in enumerate(favoriten_sortiert, 2):
    rating = eintrag.get("lb_rating")
    farbe  = bewertungsfarbe(rating)
    fill   = PatternFill("solid", fgColor=farbe) if farbe else None

    for col, (key, _) in enumerate(SPALTEN, 1):
        wert = eintrag.get(key, "")
        if key == "lb_rating" and wert:
            wert = round(float(wert), 2)
        elif key == "lb_votes" and wert:
            wert = int(wert)

        zelle = ws2.cell(row=zeile_nr, column=col, value=wert)
        zelle.font      = Font(name="Arial", size=9, bold=(key == "lb_rating"))
        zelle.alignment = Alignment(vertical="center", wrap_text=(key == "lb_desc"))
        zelle.border    = rahmen
        if fill:
            zelle.fill = fill
        if key == "lb_rating":
            zelle.number_format = "0.00"
        if key == "lb_votes":
            zelle.number_format = "#,##0"

    zeichen = len(str(eintrag.get("lb_desc") or ""))
    zeilen  = max(1, -(-zeichen // 45))
    ws2.row_dimensions[zeile_nr].height = zeilen * 15 + 6

# Spaltenbreiten
for col, (key, anzeigename) in enumerate(SPALTEN, 1):
    max_breite = len(anzeigename)
    for zeile_nr in range(2, len(favoriten_sortiert) + 2):
        wert = ws2.cell(row=zeile_nr, column=col).value
        if wert:
            max_breite = max(max_breite, min(len(str(wert)), 80))
    ws2.column_dimensions[get_column_letter(col)].width = max_breite + 2

ws2.auto_filter.ref = f"A1:{get_column_letter(len(SPALTEN))}1"
ws2.freeze_panes    = "A2"

print(f"✓ Tabellenblatt 'Lieblingsregisseure' mit {len(favoriten_sortiert)} Filmen erstellt")


# ── Speichern ───────────────────────────────────────────────────────────
dateiname = f"arte_letterboxd_{datetime.now().strftime('%Y%m%d')}.xlsx"
wb.save(dateiname)
print(f"✓ Gespeichert: {dateiname}")
print(f"  {len(sortiert)} Filme   |   "
      f"{sum(1 for e in sortiert if (e.get('lb_rating') or 0) > 4.0)} sehr gut   |   "
      f"{sum(1 for e in sortiert if 3.66 <= (e.get('lb_rating') or 0) <= 4.0)} gut   |   "
      f"{sum(1 for e in sortiert if 3.4 <= (e.get('lb_rating') or 0) < 3.66)} okay")

aktualisiere_datenbank(ergebnisse)
