# ecocardiocane

#  Descrizione
Questo script Python automatizza l’estrazione di dati da file Excel (.xls/.xlsx) in cartella sources/ e ne aggiorna un database centrale (databasecopy.xlsx), applicando trasformazioni sui valori letti.


# Prerequisiti
Python 3.7+

`pip install pandas openpyxl xlrd`


# Configurazione
/sources = inserisci i file .xls degli ECD cardiaci

/database/databasecopy.xlsx = file .xlsx in cui verranno inseriti i valori dallo script


# Esecuzione
Da terminale:
`python ecocardiocane.py`

Oppure, in un IDE, esegui direttamente il file.


# Utilizzo
Una volta eseguito lo script, copia le righe necessarie e incolla nel database di destinazione

