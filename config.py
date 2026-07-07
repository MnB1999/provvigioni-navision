"""Configurazione dell'applicazione. Unico punto in cui si cambiano percorsi e parametri."""
from pathlib import Path

# Cartella base di lavoro. Cambiare qui se serve un percorso diverso.
BASE = Path(r"C:\Provvigioni")

INBOX = BASE / "inbox"        # qui l'operatore salva gli export Excel di Navision
SCARTATI = BASE / "scartati"  # file non validi o duplicati (il motivo compare a video)
OUTPUT = BASE / "output"      # una sottocartella per ogni elaborazione completata

POLL_MS = 2000  # ogni quanti millisecondi l'app controlla la cartella inbox
