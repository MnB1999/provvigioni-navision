# provvigioni-navision
Il programma prende in input uno o più file Excel generati da Navision e li trasforma in un formato più leggibile e adatto al calcolo delle provvigioni.

Durante l'elaborazione:

mantiene esclusivamente il foglio "Visualizzazione - Fatture vend";

elimina tutti gli altri fogli presenti nel file;

conserva solamente le colonne con tipologia prodotto, quantità e importo fatturato;

rimuove le righe prive di importo (fatturate a zero);

ordina automaticamente le informazioni raggruppandole per cliente e fattura;

genera un file Excel finale con una tabella ottimizzata per il calcolo provvigionale.

Compatibilità

Il programma funziona sia con:

file Excel scaricati direttamente da Navision tramite Citrix;

file .xlsx copiati o spostati manualmente nella cartella INBOX.
