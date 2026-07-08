# provvigioni-navision
Piccolissimo tool Python che prende in input dei file xlsx salvati su disco dall'ambiente virtuale di Citrix e li converte 
in un file più leggibile tagliando tutte le informazioni non necessarie al fine del calcolo provvigionale dei nostri agenti
nello specifico il tool elimina ogni pagina dal file excel in input, tranne che per la pagina "Visualizzazione - Fatture vend"
che è il nome della pagina dove si trovano le informazioni necessarie al calcolo provvigionale
questo perché citrix genera gli excel tutti con lo stesso layout
dopodiché, da quella pagina conserva le colonne A, D, E, F e J che non abbiano campi vuoti
ordina tutte le righe suddividendo "clienti" e "fatture" e infine genera una tabella col calcolo provvigionale
Funziona sia con download diretti da Navision che con qualsiasi movimento dal pc alla tabella INBOX
