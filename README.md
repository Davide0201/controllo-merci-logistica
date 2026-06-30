controllo-merci-logistica
Script Python per il controllo automatico di merci in entrata e gestione DDT.
Cosa fa:
Lo script analizza un elenco di merci ricevute (in formato Excel) e genera automaticamente un report che evidenzia:
	•	Duplicati: righe con lo stesso codice DDT ripetuto per errore
	•	Dati mancanti: righe con campi obbligatori vuoti (es. quantità o fornitore non inseriti)
	•	Riepilogo per fornitore: conteggio degli articoli ricevuti da ciascun fornitore
Il risultato è un file Excel con tre fogli separati, uno per ciascuna verifica.
Perché l’ho fatto?
Lavoro come operatore logistico presso CAL Srl, dove mi occupo del controllo di conformità tra merce ricevuta e documentazione di trasporto. Ho creato questo script per automatizzare un controllo che normalmente richiede tempo se fatto a mano riga per riga, applicando le mie competenze di programmazione a un problema reale che incontro nel mio lavoro.
Tecnologie usate:
	•	Python 3
	•	pandas – per l’elaborazione dei dati
	•	openpyxl – per la lettura e scrittura dei file Excel
Come funziona
	1.	Lo script legge un file Excel con colonne: Codice DDT, Fornitore, Descrizione Materiale, Quantità, Data Ricevimento
	2.	Controlla automaticamente duplicati e dati mancanti
	3.	Calcola un riepilogo del numero di articoli per fornitore
	4.	Genera un report Excel finale con tutti i risultati, pronto da consultare
Esempio di output
Eseguendo lo script su un set di dati di prova, il report mostra ad esempio:
	•	2 righe duplicate trovate
	•	2 righe con dati mancanti
	•	2 fornitori distinti
Possibili sviluppi futuri:
	•	Aggiungere un controllo automatico sulle date di ricevimento
	•	Esportare il report anche in formato PDF
	•	Creare una semplice interfaccia per caricare il file senza modificare il codice
