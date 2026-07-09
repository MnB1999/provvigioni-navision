"""
Attribuzione di ogni riga fatturata a una voce della tabella di calcolo provvigionale.

Un nuovo gas da conteggiare va aggiunto alle regole prima di avviare l'applicazione
Il primo match vince quindi l'ordine conta

"""
import re

# Voci della tabella provvigioni
VOCI = (
    "Gas tecnici",
    "Elio forza maggiore",
    "Gas refrigeranti, smaltimento refrigeranti e propano",
    "Gas alimentari e puri",
    "Adr",
    "Trasporto",
    "Aqs e addizionali vari",
    "Materiali",  # non si popola mai in automatico perché le fatture dei materiali non le generiamo tramite Excel da Navision
    "Nolo e RMR",
)

# Voce speciale per le righe di adeguamento FGAS: viene riconosciuta (nessun errore
# di regola inesistente) ma, non essendo elencata in VOCI, non entra nella tabella
# provvigioni. Viene invece sommata a parte da genera_excel.py.
VOCE_ADEGUAMENTO_FGAS = "Adeguamento fgas"

# (voce, prefissi, parole): la descrizione (in maiuscolo) appartiene alla voce solo
# se inizia con uno dei prefissi oppure contiene una delle parole intere.
# Una voce può comparire più volte per gestire le eccezioni con l'ordine.
REGOLE = (
    (VOCE_ADEGUAMENTO_FGAS, (), ("ADEGUAMENTO",)),
    ("Aqs e addizionali vari", ("ADDIZIONALE", "ADEMPIMENTI", "CONTRIBUTO ENERGIA"), ()),
    ("Adr", (), ("ADR",)),
    ("Trasporto", ("TRASPORTO",), ()),
    ("Nolo e RMR", ("RMR", "RIMBORSO MANCATA RESA"), ("NOLO",)),
    ("Elio forza maggiore", (), ("FORZA MAGGIORE",)),
    ("Gas refrigeranti, smaltimento refrigeranti e propano", ("GAS REFRIGERANTE", "PROPANO", "SMALTIMENTO", "744"), ()),
    ("Gas tecnici", ("AZOTO 5.0", "ELIO 5.0"), ()),  # eccezione: prima dei gradi di purezza
    ("Gas alimentari e puri", (), ("TRESARIS", "5.0", "5.5", "6.0")),
    ("Gas tecnici", ("OSSIGENO", "ACETILENE", "CORGON", "CRONIGON", "FORMIN", "ANIDRIDE", "CO2"), ()),
)


class VoceNonClassificabile(ValueError):
    """La descrizione non corrisponde ad alcuna regola: va aggiunta in REGOLE."""


def voce_per(descrizione: str) -> str:
    """Restituisce la voce della tabella a cui appartiene la descrizione."""
    testo = descrizione.upper()
    for voce, prefissi, parole in REGOLE:
        if any(testo.startswith(p) for p in prefissi):
            return voce
        if any(re.search(rf"\b{re.escape(parola)}\b", testo) for parola in parole):
            return voce
    raise VoceNonClassificabile(descrizione)
