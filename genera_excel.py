"""Scrittura dei file provvigioni: un file Excel per ogni agente.

Struttura di ogni file (replica il formato del file di riferimento):
- per ogni cliente (in ordine alfabetico), le sue fatture in ordine di numero;
- la prima fattura del cliente ha intestazione "CLIENTE - Fattura nr NUMERO",
  le successive "ft NUMERO";
- sotto ogni intestazione, le righe della fattura su 5 colonne:
  Tipo | Descrizione | Quantità | Unità di misura | Importo;
- una riga vuota separa ogni fattura.
"""
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from fattura import Fattura

CARATTERE_TITOLO = Font(name="Arial", size=12, bold=True)
CARATTERE_DATI = Font(name="Verdana", size=10)
LARGHEZZE_COLONNE = {"A": 12, "B": 52, "C": 9, "D": 9, "E": 13}


def scrivi_provvigioni(fatture: list[Fattura], cartella: Path) -> list[Path]:
    """Raggruppa le fatture per agente e scrive un file per agente nella cartella. Restituisce i percorsi creati."""
    if not fatture:
        raise ValueError("nessuna fattura da elaborare")
    cartella.mkdir(parents=True, exist_ok=True)

    per_agente: dict[str, list[Fattura]] = defaultdict(list)
    for fattura in fatture:
        per_agente[fattura.codice_agente].append(fattura)

    return [
        _scrivi_file_agente(agente, per_agente[agente], cartella)
        for agente in sorted(per_agente)
    ]


def _scrivi_file_agente(agente: str, fatture: list[Fattura], cartella: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]:*?/\\]", "_", agente)[:31]
    for colonna, larghezza in LARGHEZZE_COLONNE.items():
        ws.column_dimensions[colonna].width = larghezza

    per_cliente: dict[str, list[Fattura]] = defaultdict(list)
    for fattura in fatture:
        per_cliente[fattura.cliente].append(fattura)

    for cliente in sorted(per_cliente):
        in_ordine = sorted(per_cliente[cliente], key=lambda f: f.numero)
        for posizione, fattura in enumerate(in_ordine):
            if posizione == 0:
                titolo = f"{cliente} - Fattura nr {fattura.numero}"
            else:
                titolo = f"ft {fattura.numero}"
            ws.append([titolo])
            ws.cell(row=ws.max_row, column=1).font = CARATTERE_TITOLO
            for riga in fattura.righe:
                ws.append([riga.tipo, riga.descrizione, riga.quantita, riga.unita_misura, riga.importo])
                for cella in ws[ws.max_row]:
                    cella.font = CARATTERE_DATI
            ws.append([None])  # riga vuota di separazione

    nome_sicuro = re.sub(r"[^\w-]", "_", agente)
    percorso = cartella / f"Provvigioni_{nome_sicuro}.xlsx"
    wb.save(percorso)
    return percorso
