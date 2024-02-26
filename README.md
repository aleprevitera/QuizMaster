# QuizMaster

## Version 1.1.0

Introduzione dell'interfaccia grafica utente (GUI).

QuizMaster è un'applicazione desktop sviluppata per facilitare la creazione e la gestione di quiz con i relativi correttori. Con l'introduzione della versione 1.1.0, QuizMaster ora include un'interfaccia grafica utente, rendendo il processo ancora più intuitivo e accessibile.

### Funzionalità

- **Selezione dell'archivio di domande**: Permette di selezionare un file Excel contenente le domande del quiz.
- **Configurazione del quiz**: Permette di inserire i dettagli del quiz, come materia, CDL, anno, sezione, data, e numero di domande.
- **Generazione dei PDF**: Genera un PDF contenente le domande selezionate casualmente dal file Excel specificato e un PDF contente il correttore del quiz.

### Archivi

Gli archivi devono essere in formato **.xlsx** e strutturati a 6 colonne disposte così:

DOMANDA | RISPOSTA CORRETTA | Testo2 | Testo3 | Testo4 | Testo5

### Dipendenze

- Python 3
- Tkinter per l'interfaccia grafica
- Pandas per la manipolazione dei dati
- ReportLab per la generazione di file PDF

### Installazione delle dipendenze

Per installare le dipendenze necessarie, apri il terminale e esegui:

```bash
pip install pandas reportlab openpyxl tk
```
### Esecuzione

Per eseguire l'applicazione, naviga alla directory del progetto e esegui:

```bash
python path_to_your_script.py
```
Sostituisci path_to_your_script.py con il percorso effettivo del tuo script. Assicurati che tutte le dipendenze siano state installate correttamente prima di eseguire lo script.

### Sviluppo

Questo progetto è stato sviluppato con l'obiettivo di fornire uno strumento utile per la creazione rapida e efficiente di materiali per quiz. Feedback e contributi sono sempre ben accetti per migliorare ulteriormente l'applicazione.

