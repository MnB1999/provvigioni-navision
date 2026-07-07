"""Applicazione Provvigioni Navision.

All'avvio crea le cartelle di lavoro e controlla la inbox ogni 2 secondi.
Ogni export di fattura salvato nella inbox viene letto, validato e mostrato in tabella.
I file non validi o duplicati finiscono in "scartati", con il motivo a video.
Il pulsante "Genera file provvigioni" scrive un file Excel per agente e archivia
gli export originali nella stessa cartella di output, sotto "originali".
"""
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import config
from fattura import Fattura, FatturaNonValida, leggi_fattura
from genera_excel import scrivi_provvigioni


class Applicazione:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.in_attesa: dict[str, tuple[int, float]] = {}  # nome file -> (dimensione, mtime) all'ultimo controllo
        self.gia_letti: set[str] = set()                   # nomi file già acquisiti (restano in inbox fino alla generazione)
        self.fatture: dict[str, Fattura] = {}              # numero fattura -> Fattura
        self.file_per_numero: dict[str, Path] = {}         # numero fattura -> file in inbox
        self._costruisci_interfaccia()
        self.root.after(config.POLL_MS, self._controlla_inbox)

    # ------------------------------------------------------------ interfaccia

    def _costruisci_interfaccia(self) -> None:
        self.root.title("Provvigioni Navision")
        self.root.geometry("1000x560")

        istruzioni = (
            f"1. In Navision apri la fattura in Excel e salva il file nel formato .XLXS in:  {config.INBOX}\n"
            "2. Ripeti per tutte le fatture del trimestre, di tutti gli agenti, in qualsiasi ordine.\n"
            "3. Le fatture acquisite compaiono qui sotto. Al termine premi \u00abGenera file provvigioni\u00bb."
        )
        tk.Label(self.root, text=istruzioni, justify="left", anchor="w", padx=10, pady=8).pack(fill="x")

        colonne = [
            ("numero", "Fattura", 110),
            ("data", "Data", 90),
            ("agente", "Agente", 80),
            ("cliente", "Cliente", 380),
            ("stato", "Stato", 280),
        ]
        self.tabella = ttk.Treeview(self.root, columns=[c[0] for c in colonne], show="headings")
        for nome, testo, larghezza in colonne:
            self.tabella.heading(nome, text=testo)
            self.tabella.column(nome, width=larghezza, anchor="w")
        self.tabella.tag_configure("errore", foreground="red")
        self.tabella.pack(fill="both", expand=True, padx=10)

        barra = tk.Frame(self.root)
        barra.pack(fill="x", padx=10, pady=8)
        self.contatore = tk.Label(barra, text="Fatture acquisite: 0")
        self.contatore.pack(side="left")
        self.pulsante = tk.Button(barra, text="Genera file provvigioni", state="disabled", command=self._genera)
        self.pulsante.pack(side="right")

    # -------------------------------------------------------- acquisizione

    def _controlla_inbox(self) -> None:
        """Un file viene letto solo quando dimensione e data di modifica restano
        uguali tra due controlli consecutivi: così non si legge un file che
        Citrix sta ancora copiando."""
        try:
            for percorso in sorted(config.INBOX.glob("*.xlsx")):
                nome = percorso.name
                if nome in self.gia_letti or nome.startswith("~$"):
                    continue
                stato = percorso.stat()
                firma = (stato.st_size, stato.st_mtime)
                if self.in_attesa.get(nome) == firma:
                    self._acquisisci(percorso)
                else:
                    self.in_attesa[nome] = firma
        finally:
            self.root.after(config.POLL_MS, self._controlla_inbox)

    def _acquisisci(self, percorso: Path) -> None:
        try:
            fattura = leggi_fattura(percorso)
        except FatturaNonValida as errore:
            self._scarta(percorso, str(errore))
            return
        if fattura.numero in self.fatture:
            self._scarta(percorso, f"duplicato della fattura {fattura.numero}")
            return
        self.gia_letti.add(percorso.name)
        self.fatture[fattura.numero] = fattura
        self.file_per_numero[fattura.numero] = percorso
        self.tabella.insert("", "end", values=(
            fattura.numero,
            fattura.data_documento.strftime("%d/%m/%Y"),
            fattura.codice_agente,
            fattura.cliente,
            "acquisita",
        ))
        self._aggiorna_contatore()

    def _scarta(self, percorso: Path, motivo: str) -> None:
        config.SCARTATI.mkdir(parents=True, exist_ok=True)
        destinazione = config.SCARTATI / f"{datetime.now():%Y%m%d_%H%M%S}_{percorso.name}"
        shutil.move(str(percorso), destinazione)
        self.in_attesa.pop(percorso.name, None)
        self.tabella.insert(
            "", "end",
            values=("—", "", "", percorso.name, f"SCARTATA: {motivo}"),
            tags=("errore",),
        )

    def _aggiorna_contatore(self) -> None:
        agenti = {f.codice_agente for f in self.fatture.values()}
        self.contatore.config(text=f"Fatture acquisite: {len(self.fatture)}  •  Agenti: {len(agenti)}")
        self.pulsante.config(state="normal" if self.fatture else "disabled")

    # -------------------------------------------------------- generazione

    def _genera(self) -> None:
        cartella = config.OUTPUT / datetime.now().strftime("%Y-%m-%d_%H%M%S")
        creati = scrivi_provvigioni(list(self.fatture.values()), cartella)

        originali = cartella / "originali"
        originali.mkdir()
        for numero, percorso in self.file_per_numero.items():
            nome_sicuro = re.sub(r"[^\w-]", "_", numero)
            shutil.move(str(percorso), originali / f"{nome_sicuro}.xlsx")

        quante = len(self.fatture)
        self.fatture.clear()
        self.file_per_numero.clear()
        self.gia_letti.clear()
        self.in_attesa.clear()
        for riga in self.tabella.get_children():
            self.tabella.delete(riga)
        self._aggiorna_contatore()

        messagebox.showinfo(
            "Elaborazione completata",
            f"Elaborate {quante} fatture.\nCreati {len(creati)} file provvigioni in:\n{cartella}",
        )
        os.startfile(cartella)


def main() -> None:
    for cartella in (config.INBOX, config.SCARTATI, config.OUTPUT):
        cartella.mkdir(parents=True, exist_ok=True)
    root = tk.Tk()
    Applicazione(root)
    root.mainloop()


if __name__ == "__main__":
    main()
