#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mubi.py — MUBI-Integration für den Mediathek Film Tracker

Holt von MUBI über die öffentliche (inoffizielle) API:
  - Watchlist ("wishes")  → /v4/users/{id}/wishes
  - eigene Ratings        → /v3/users/{id}/ratings

Daraus entstehen drei Features:
  1. Watchlist-Abgleich : welche Watchlist-Filme laufen gerade in den Mediatheken?
  2. Eigene MUBI-Ratings : als zusätzliche Spalte 'mubi_rating'
  3. Lieblingsregisseure : Regie mit >= 2 Filmen, die du mit 4 oder 5 Sternen
                           bewertet hast → lieblingsregisseure_auto.json

Standalone-Test (ohne Mediathek-Daten):
    python mubi.py
"""

# %% ── Imports und Konfiguration ─────────────────────────────────────────────
import os
import json
import time
import requests

from config import MUBI_USER_ID   # neu in config.py (deine MUBI-User-ID)

# Regeln für die automatische Lieblingsregisseur-Erkennung
MIN_STERNE          = 4   # nur Filme mit >= 4 Sternen zählen als "geliebt"
MIN_FILME_PRO_REGIE = 2   # Regie muss >= so viele solcher Filme haben

# Datei für die automatisch erkannten Regisseur:innen (getrennt von config.py,
# damit deine handgepflegte Liste nie automatisch überschrieben wird)
AUTO_REGIE_PFAD = "lieblingsregisseure_auto.json"

# MUBI verlangt diese Header — ohne 'Client'/'Client-Country' kommt 406/403.
# Wichtig: KEIN Accept-Language setzen, das löste in Tests einen 406 aus.
MUBI_HEADER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Client": "web",
    "Client-Country": "DE",
}

PER_PAGE = 24   # Einträge pro API-Seite


# %% ── Schritt 1: Generischer paginierter Abruf ──────────────────────────────
def _hole_alle_seiten(url, daten_schluessel, max_seiten=50):
    """
    Ruft einen paginierten MUBI-Endpunkt komplett ab.

    MUBI paginiert über den Parameter 'before': Jede Antwort enthält
    meta.next_cursor. Diesen Wert geben wir bei der nächsten Anfrage als
    'before' mit → liefert die nächste Portion. Ist next_cursor None oder
    ändert er sich nicht mehr, sind wir am Ende.

    daten_schluessel: Schlüssel im JSON ('wishes' bzw. 'ratings').
    max_seiten:       Sicherheitslimit gegen Endlosschleifen.
    """
    alle_eintraege = []
    before  = None          # erste Anfrage ohne 'before'
    runde   = 0

    while runde < max_seiten:
        params = {"per_page": PER_PAGE}
        if before is not None:
            params["before"] = before

        antwort = requests.get(url, headers=MUBI_HEADER, params=params, timeout=15)
        antwort.raise_for_status()
        daten = antwort.json()

        eintraege = daten.get(daten_schluessel, [])
        if not eintraege:
            break
        alle_eintraege.extend(eintraege)

        # Nächsten Cursor lesen und als 'before' für die nächste Runde merken
        meta             = daten.get("meta", {})
        naechster_cursor = meta.get("next_cursor")

        # Ende erreicht: kein Cursor mehr ODER Cursor wiederholt sich
        if naechster_cursor is None or naechster_cursor == before:
            break

        before = naechster_cursor
        runde += 1
        time.sleep(0.4)   # höflich zur API

    return alle_eintraege


def _extrahiere_regie(film):
    """
    Liest die Regie aus einem MUBI-'film'-Objekt.

    MUBI liefert die Regie üblicherweise als Liste unter 'directors',
    je Eintrag mit 'name'. Falls die Struktur abweicht, fängt diese
    Funktion das defensiv ab und gibt notfalls '' zurück.
    """
    regisseure = film.get("directors", [])

    namen = []
    for person in regisseure:
        if isinstance(person, dict):
            name = person.get("name", "")
        else:
            name = str(person)
        if name:
            namen.append(name)

    return ", ".join(namen)


# %% ── Schritt 2: Watchlist holen ───────────────────────────────────────────
def hole_watchlist(user_id=MUBI_USER_ID):
    """
    Holt die komplette Watchlist ('wishes') und gibt eine Liste
    vereinfachter Film-Dicts zurück (Titel, Originaltitel, Jahr, Regie).
    """
    url           = f"https://api.mubi.com/v4/users/{user_id}/wishes"
    rohe_wuensche = _hole_alle_seiten(url, "wishes")

    watchlist = []
    for wunsch in rohe_wuensche:
        film = wunsch.get("film", {})
        eintrag = {
            "mubi_id":        film.get("id"),
            "titel":          film.get("title", ""),
            "original_titel": film.get("original_title", ""),
            "jahr":           film.get("year", ""),
            "regie":          _extrahiere_regie(film),
        }
        watchlist.append(eintrag)

    print(f"✓ MUBI-Watchlist: {len(watchlist)} Filme geladen")
    return watchlist


# %% ── Schritt 3: Eigene Ratings holen ──────────────────────────────────────
def hole_ratings(user_id=MUBI_USER_ID):
    """
    Holt alle eigenen Ratings und gibt zwei Dinge zurück:

      ratings_nach_titel : dict {titel_klein: sterne}  (für den Mediathek-Abgleich;
                           sowohl deutscher als auch Originaltitel als Schlüssel)
      bewertete_filme    : Liste {titel, regie, sterne} (für die Regie-Auswertung)
    """
    url          = f"https://api.mubi.com/v3/users/{user_id}/ratings"
    rohe_ratings = _hole_alle_seiten(url, "ratings")

    ratings_nach_titel = {}
    bewertete_filme    = []

    for rating in rohe_ratings:
        sterne = rating.get("overall")
        film   = rating.get("film", {})

        titel          = film.get("title", "")
        original_titel = film.get("original_title", "")
        regie          = _extrahiere_regie(film)

        # Beide Titelvarianten als Suchschlüssel ablegen → robusterer Abgleich
        if titel:
            ratings_nach_titel[titel.lower().strip()] = sterne
        if original_titel:
            ratings_nach_titel[original_titel.lower().strip()] = sterne

        bewertete_filme.append({
            "titel":  titel,
            "regie":  regie,
            "sterne": sterne,
        })

    print(f"✓ MUBI-Ratings: {len(bewertete_filme)} bewertete Filme geladen")
    return ratings_nach_titel, bewertete_filme


# %% ── Schritt 4: Lieblingsregisseure automatisch bestimmen ──────────────────
def aktualisiere_lieblingsregisseure(bewertete_filme):
    """
    Bestimmt Lieblingsregisseur:innen nach der Regel:
      mindestens MIN_FILME_PRO_REGIE Filme mit >= MIN_STERNE Sternen.

    Schreibt das Ergebnis nach AUTO_REGIE_PFAD (JSON) und gibt die
    Namensliste zurück.
    """
    # Pro Regie die hoch bewerteten Filme zählen
    zaehler_pro_regie = {}

    for film in bewertete_filme:
        sterne = film.get("sterne") or 0
        if sterne < MIN_STERNE:
            continue

        regie_string = film.get("regie", "")
        if not regie_string:
            continue

        # Ein Film kann mehrere Regisseur:innen haben → einzeln zählen
        einzelne_namen = [n.strip() for n in regie_string.split(",") if n.strip()]
        for name in einzelne_namen:
            if name not in zaehler_pro_regie:
                zaehler_pro_regie[name] = 0
            zaehler_pro_regie[name] += 1

    # Nur Regie mit genügend Filmen behalten
    lieblingsregisseure = []
    for name, anzahl in zaehler_pro_regie.items():
        if anzahl >= MIN_FILME_PRO_REGIE:
            lieblingsregisseure.append(name)
    lieblingsregisseure.sort()

    # In JSON speichern (wird auf GitHub committet → bleibt erhalten)
    with open(AUTO_REGIE_PFAD, "w", encoding="utf-8") as datei:
        json.dump(lieblingsregisseure, datei, ensure_ascii=False, indent=2)

    print(f"✓ {len(lieblingsregisseure)} automatische Lieblingsregisseure "
          f"(>= {MIN_FILME_PRO_REGIE} Filme mit >= {MIN_STERNE}★) → {AUTO_REGIE_PFAD}")
    return lieblingsregisseure


def lade_auto_regisseure(pfad=AUTO_REGIE_PFAD):
    """
    Liest die automatisch erkannten Lieblingsregisseure aus der JSON.
    Fehlt die Datei, kommt eine leere Liste zurück.
    """
    if not os.path.exists(pfad):
        return []
    with open(pfad, "r", encoding="utf-8") as datei:
        return json.load(datei)


# %% ── Schritt 5: Watchlist gegen Mediathek-Ergebnisse abgleichen ────────────
def finde_watchlist_treffer(mediathek_ergebnisse, watchlist):
    """
    Prüft, welche Watchlist-Filme aktuell in den Mediatheken laufen.
    Abgleich über Titel (deutsch UND original), case-insensitiv,
    BEWERTUNGSUNABHÄNGIG.

    Gibt die passenden Mediathek-Einträge zurück.
    """
    # Alle Watchlist-Titel (beide Varianten) in ein Set für schnellen Vergleich
    watchlist_titel = set()
    for film in watchlist:
        if film.get("titel"):
            watchlist_titel.add(film["titel"].lower().strip())
        if film.get("original_titel"):
            watchlist_titel.add(film["original_titel"].lower().strip())

    treffer = []
    for eintrag in mediathek_ergebnisse:
        lb_titel    = (eintrag.get("lb_title") or "").lower().strip()
        quell_titel = (eintrag.get("quell_titel") or "").lower().strip()

        if lb_titel in watchlist_titel or quell_titel in watchlist_titel:
            treffer.append(eintrag)

    print(f"✓ {len(treffer)} Watchlist-Filme aktuell in den Mediatheken")
    return treffer


# %% ── Schritt 6: Eigene MUBI-Ratings an Mediathek-Ergebnisse anhängen ───────
def haenge_mubi_ratings_an(mediathek_ergebnisse, ratings_nach_titel):
    """
    Ergänzt jeden Mediathek-Eintrag um das eigene MUBI-Rating
    (Schlüssel 'mubi_rating'), sofern ein Titel-Match existiert.
    Verändert die Liste in-place und gibt sie zurück.
    """
    for eintrag in mediathek_ergebnisse:
        lb_titel    = (eintrag.get("lb_title") or "").lower().strip()
        quell_titel = (eintrag.get("quell_titel") or "").lower().strip()

        mubi_rating = ratings_nach_titel.get(lb_titel)
        if mubi_rating is None:
            mubi_rating = ratings_nach_titel.get(quell_titel)

        eintrag["mubi_rating"] = mubi_rating   # None, wenn kein Match
    return mediathek_ergebnisse


# %% ── Orchestrierung: alles in einem Aufruf ────────────────────────────────
def reichere_mit_mubi_an(mediathek_ergebnisse):
    """
    Führt die komplette MUBI-Integration aus und gibt zurück:
      - mediathek_ergebnisse (in-place um 'mubi_rating' ergänzt)
      - watchlist_treffer    (Liste für das separate Excel-Blatt)
      - auto_regisseure      (aktualisierte Namensliste)

    Der Aufrufer (main.py) fängt Exceptions ab, damit ein MUBI-Ausfall
    die Hauptpipeline nicht stoppt.
    """
    watchlist                       = hole_watchlist()
    ratings_nach_titel, bewertete   = hole_ratings()

    auto_regisseure   = aktualisiere_lieblingsregisseure(bewertete)
    watchlist_treffer = finde_watchlist_treffer(mediathek_ergebnisse, watchlist)
    haenge_mubi_ratings_an(mediathek_ergebnisse, ratings_nach_titel)

    return mediathek_ergebnisse, watchlist_treffer, auto_regisseure


# %% ── Selbsttest (direkt ausführbar) ───────────────────────────────────────
if __name__ == "__main__":
    print("=== MUBI-Modul Selbsttest ===\n")

    watchlist                     = hole_watchlist()
    ratings_nach_titel, bewertete = hole_ratings()
    auto                          = aktualisiere_lieblingsregisseure(bewertete)

    # WICHTIG: prüft, ob die Regie korrekt ausgelesen wird.
    # Falls hier "KEINE Regie gefunden" steht, stimmt der Feldname in
    # _extrahiere_regie() nicht — dann melden, wir passen ihn an.
    print("\n── Kontrolle: werden Regie-Namen gefunden? ──")
    beispiele_mit_regie = []
    for film in bewertete:
        if film["regie"]:
            beispiele_mit_regie.append(film)
        if len(beispiele_mit_regie) >= 5:
            break

    if beispiele_mit_regie:
        for film in beispiele_mit_regie:
            print(f"  {film['sterne']}★  {film['regie']:30s}  {film['titel']}")
    else:
        print("  ⚠ KEINE Regie gefunden — Feldname in _extrahiere_regie() prüfen!")

    print(f"\nAutomatische Lieblingsregisseure: {auto}")
    if watchlist:
        print(f"Beispiel Watchlist-Film: {watchlist[0]}")


# %% ── Diagnose: paginiert der wishes-Endpunkt korrekt? ──────────────────────
import requests, json

MUBI_USER_ID = 19068609
url = f"https://api.mubi.com/v4/users/{MUBI_USER_ID}/wishes"
header = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Client": "web", "Client-Country": "DE",
}

