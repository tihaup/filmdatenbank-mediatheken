# Projektregeln für Mediathek-Tracker

## Coding-Stil (NICHT verändern)
- Variablennamen auf Deutsch: rohdaten, gefilterte_filme, ergebnisse
- Kommentare auf Deutsch
- Intermediate Variablen statt Method Chaining
- For-Loops statt List Comprehensions wo möglich
- Funktionsnamen auf Deutsch: lade_mediathek_filme, waehle_beste_version

## Was Claude Code NICHT tun darf
- Keine bestehenden Funktionen umbenennen ohne explizite Aufforderung
- Keine Variablennamen ändern ohne explizite Aufforderung
- Keine Refaktorisierung von funktionierendem Code
- Nie mehr als eine Sache auf einmal ändern
- Nicht 'verbessern' was nicht im Prompt steht

## Arbeitsweise
- Nach jeder Änderung: zeige mir NUR die geänderten Stellen
- Erkläre kurz WAS du geändert hast und WARUM
- Wenn unklar: frage nach, handle nicht

## Projektstruktur
- notebooks/ : Jupyter Notebooks
- data/      : CSV-Dateien (nie löschen)
- output/    : Excel-Ausgaben
