# Sender und deren Suchthemen
SENDER = {
    "ARTE.DE": "filme",
    "ZDF": "film",
    "ARD": "film",
    "3sat": "#Spielfilm",
    "BR": "Filme"
}

# Lieblingsregisseure werden automatisch aus den MUBI-Ratings bestimmt
# (siehe mubi.py). Hier nur noch manuelle Ergänzungen eintragen, falls gewünscht.
LIEBLINGSREGISSEURE = []

# Bewertungsschwellen
FARBE_SEHR_GUT = "A8D5A2"   # > 4.0
FARBE_GUT      = "D6EFD8"   # 3.66 - 4.0
FARBE_OK       = "FFF3CD"   # 3.4 - 3.65

# Mindestlaufzeit in Sekunden
MIN_DURATION_SEC = 3600

# MUBI-Integration
MUBI_USER_ID = 19068609