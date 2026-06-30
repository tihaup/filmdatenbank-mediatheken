# 🎬 Mediathek Film Tracker

Dieses Projekt durchsucht jede Woche automatisch die Mediatheken der öffentlich-rechtlichen Sender (Arte, ZDF, ARD, 3sat) nach Spielfilmen und reichert sie mit Bewertungen von **Letterboxd** an. Zusätzlich gleicht es die Filme mit deinem **MUBI**-Konto ab: Es zeigt dir, welche Filme von deiner MUBI-Watchlist gerade verfügbar sind, blendet deine eigenen MUBI-Bewertungen ein und erkennt automatisch deine Lieblingsregisseur:innen anhand deiner am besten bewerteten Filme. Das Ergebnis ist eine übersichtliche, farbcodierte Excel-Tabelle und eine direkt im Browser lesbare CSV-Liste – einmal pro Woche frisch erzeugt, ganz von allein.

Technisch läuft alles über einen kostenlosen Automatik-Dienst von GitHub (GitHub Actions). Du musst nichts auf deinem Computer installieren und nichts laufen lassen – nach der einmaligen Einrichtung erledigt sich alles selbst.

---

## So nutzt du den Tracker mit deinem eigenen MUBI-Konto

Diese Anleitung ist für alle gedacht – auch ohne Programmierkenntnisse. Du brauchst nur ein GitHub-Konto (kostenlos) und ein MUBI-Konto. Plane etwa 15 Minuten ein.

> **Wichtige Voraussetzung:** Dein MUBI-Profil muss **öffentlich** sein, sonst kann der Tracker deine Watchlist und Bewertungen nicht lesen. Das stellst du in den MUBI-Einstellungen unter Privatsphäre ein.

### Schritt 1: Kostenlosen TMDb-Schlüssel besorgen

Der Tracker nutzt die Filmdatenbank TMDb, um Filme zu finden. Dafür braucht es einen kostenlosen Zugangsschlüssel (API-Key):

1. Gehe auf [themoviedb.org](https://www.themoviedb.org/) und erstelle ein kostenloses Konto.
2. Klicke oben rechts auf dein Profilbild → **Einstellungen** (Settings).
3. Wähle links im Menü **API**.
4. Beantrage einen Schlüssel („Request an API Key") – wähle den Typ **Developer**. Bei den Formularfeldern kannst du Beliebiges eintragen (z. B. „privates Hobbyprojekt").
5. Du bekommst einen **API Key (v3 auth)** – eine lange Zeichenfolge. Kopiere sie und leg sie kurz beiseite, du brauchst sie in Schritt 4.

### Schritt 2: Das Projekt in dein eigenes GitHub kopieren (forken)

1. Logge dich bei [github.com](https://github.com/) ein (oder erstelle ein kostenloses Konto).
2. Öffne dieses Repository und klicke oben rechts auf den Button **Fork**.
3. Bestätige – GitHub legt jetzt eine eigene Kopie in deinem Konto an. Mit dieser Kopie arbeitest du ab jetzt.

### Schritt 3: Deine MUBI-ID eintragen

1. Finde deine MUBI-User-ID: Öffne dein MUBI-Profil im Browser und schau in die Adresszeile. Sie sieht so aus: `mubi.com/de/users/12345678/...` – die Zahl (hier `12345678`) ist deine ID.
2. In deinem geforkten GitHub-Repo: Klicke auf die Datei **`config.py`**.
3. Klicke rechts oben auf das **Stift-Symbol** (Bearbeiten).
4. Suche die Zeile `MUBI_USER_ID = ...` und ersetze die Zahl durch deine eigene.
5. Klicke oben rechts auf **Commit changes** (Änderungen speichern).

### Schritt 4: Deinen TMDb-Schlüssel sicher hinterlegen

Damit dein Schlüssel nicht öffentlich im Code steht, speichert GitHub ihn an einem geschützten Ort („Secret"):

1. In deinem Repo: Klicke oben auf **Settings** (Einstellungen).
2. Links im Menü: **Secrets and variables** → **Actions**.
3. Klicke auf **New repository secret**.
4. Trage bei **Name** exakt `TMDB_KEY` ein (genau so geschrieben).
5. Füge bei **Secret** deinen TMDb-Schlüssel aus Schritt 1 ein.
6. Klicke **Add secret**.

### Schritt 5: Die Automatik aktivieren und den ersten Lauf starten

Bei geforkten Repos ist die Automatik aus Sicherheitsgründen zunächst deaktiviert. So schaltest du sie ein:

1. Klicke in deinem Repo oben auf den Reiter **Actions**.
2. Falls ein grüner Hinweis erscheint, klicke auf **„I understand my workflows, go ahead and enable them"**.
3. Wähle links den Workflow **Mediathek Tracker** aus.
4. Klicke rechts auf **Run workflow** → noch einmal **Run workflow**, um einen ersten Lauf von Hand zu starten.
5. Warte ein paar Minuten – der Lauf erscheint in der Liste und wird grün, wenn alles geklappt hat.

### Schritt 6: Deine Ergebnisse ansehen

Es gibt zwei Wege, an die Filmliste zu kommen:

**Schnell im Browser (empfohlen):** Öffne in deinem Repo den Ordner **`data`** und klicke auf **`aktuelle_woche.csv`**. GitHub zeigt sie direkt als Tabelle an – das ist die Liste der neuen Filme dieser Woche.

**Als formatierte Excel:** Gehe auf den Reiter **Actions**, öffne den letzten Lauf und scrolle ganz nach unten zum Abschnitt **Artifacts**. Dort liegt die farbcodierte Excel-Datei zum Herunterladen.

### Schritt 7: Zurücklehnen

Ab jetzt läuft der Tracker **jeden Freitagmorgen automatisch**. Du musst nichts weiter tun – schau einfach am Wochenende in deine `aktuelle_woche.csv` oder lade dir die neue Excel herunter.

---

## Häufige Fragen

**Der Lauf wird rot / schlägt fehl.** Meist stimmt etwas mit dem TMDb-Schlüssel nicht. Prüfe in Schritt 4, ob das Secret exakt `TMDB_KEY` heißt und der Schlüssel korrekt eingefügt wurde (ohne Anführungszeichen oder Leerzeichen).

**Meine Watchlist taucht nicht auf.** Stelle sicher, dass dein MUBI-Profil öffentlich ist (siehe Voraussetzung oben) und deine MUBI-ID in `config.py` korrekt eingetragen ist.

**Kann ich die Sender ändern?** Ja – in `config.py` im Abschnitt `SENDER`. Für die meisten ist die Voreinstellung aber genau richtig.

**Die Lieblingsregisseure stimmen nicht.** Sie werden automatisch aus deinen MUBI-Filmen mit 4 oder 5 Sternen bestimmt (ab 2 Filmen pro Regie). Je mehr du auf MUBI bewertest, desto besser wird die Liste.
