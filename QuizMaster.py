# QUIZMASTER
# v. 1.1.0
# introduzione della GUI


import tkinter as tk
from tkinter import StringVar, ttk, messagebox, filedialog
import ctypes
import os
import pandas as pd
import random
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle, TA_CENTER
from reportlab.platypus import KeepTogether

#=================PERCORSO=================
percorso = os.path.dirname(os.path.realpath(__file__))
photo = os.path.join(percorso, 'sfoglia.png')
logo = os.path.join(percorso, 'logo.ico')
#=================DPI awareness=================
ctypes.windll.shcore.SetProcessDpiAwareness(1)

#=================GEOMETRIA=================
root = tk.Tk()
root.geometry("800x1000")
root.resizable(False, False)
root.title("QuizMaster")
root.iconbitmap(logo)
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)


root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
root.rowconfigure(3, weight=1)
root.rowconfigure(4, weight=1)
root.rowconfigure(5, weight=1)
root.rowconfigure(6, weight=1)



# ==================IMMAGINE=================
icon = tk.PhotoImage(file=photo)
icon = icon.subsample(2, 2)


# ==================COLORI=================
palette = [ "#384589", "#2D3B6A", "#1D2D46", "#47366D", "#5A4485", "#2A2A2A", "#474747", "#3376CD", "#4255C8", "#606060" ]

# ==================STILI GUI=================
style = ttk.Style()
style.theme_create("agp_style", parent="alt", settings={
    "TCombobox": {
        "configure": {"fieldbackground": palette[6], "selectbackground": "#00000000",
                       "selectforeground":"white",  "font":"Montserrat 18"}
    },
    "Hover.TButton": {
        "configure": {"background": palette[8], "foreground": "white", "font":"Montserrat 13"}
    },
    "Off.TButton": {
        "configure": {"background": palette[8], "foreground": "white", "font":"Montserrat 13"}
    },
    "TButton": {
        "layout": [
            ("Button.border", {"sticky": "nswe", "children": [
                ("Button.focus", {"sticky": "nswe", "children": [
                    ("Button.padding", {"sticky": "nswe", "children": [
                        ("Button.label", {"sticky": "ns"})
                    ]})
                ]})
            ]})
        ],
        "configure": {"background": palette[7], "foreground": "white", "font":"Montserrat 13", "padding": (15, 5)}
    },
    "TLabel": {
        "configure": {"foreground": "white", "font":"Montserrat 13", "background": palette[5],
                      "padding": "5 5 5 5"}
    },
    "Off.TLabel": {
        "configure": {"foreground": palette[9], "font":"Montserrat 13", "background": palette[5],
                      "padding": "5 5 5 5"}
    },
    "Hover.TEntry": {
        "configure": {"fieldbackground": palette[8], "selectbackground": "#00000000", "foreground": "white", "padding": "10 5", "font":"Montserrat 13"}
    },
    "TEntry": {
        "configure": {"fieldbackground": palette[6], "selectbackground": "#00000000", "foreground": "white", "padding": "10 5", "font":"Montserrat 13"}
    },
     "Off.TEntry": {
        "configure": {"fieldbackground": palette[9], "selectbackground": "#00000000", "foreground": "white", "font":"Montserrat 13"}
    }
})
style.theme_use("agp_style")
root.configure(bg=palette[5])

# ==================STILI PDF=================
bold_style = ParagraphStyle(
    name='BoldStyle',
    fontName='Helvetica-Bold',  
    fontSize=11,
    leading=14,
    textColor='black',
)
normal_style = ParagraphStyle(
    name='NormalStyle',
    fontName='Helvetica', 
    fontSize=11,
    leading=14,
    textColor='black',
)

corretta_style = ParagraphStyle(
    name='CorrettaStyle',
    fontName='Helvetica-Bold', 
    fontSize=11,
    leading=14,
    textColor='black',
    backColor='yellow',
)

centered_style = ParagraphStyle(
    name='CenteredStyle',
    fontName='Helvetica-Bold',
    fontSize=16,
    leading=18,
    textColor='black',
    alignment=TA_CENTER, 
)

sub_centered_style = ParagraphStyle(
    name='SubCenteredStyle',
    fontName='Helvetica-Bold',
    fontSize=12,
    leading=18,
    textColor='black',
    alignment=TA_CENTER,
)

spazio_style = ParagraphStyle(
    name='SpazioStyle',
    textColor='white',
)

# ==================FUNZIONi=================
def funzione():
    print(archivio.get())

def on_enter(event):
    button.configure(style="Hover.TButton")
def on_enter2(event):
    button2.configure(style="Hover.TButton")


def on_leave(event):
    button.configure(style="TButton")
def on_leave2(event):
    button2.configure(style="TButton")

def attiva(event):
    intestazione_testo_entry.configure(state="normal", style="TEntry")
    cdl_entry.configure(state="normal", style="TEntry")
    anno_entry.configure(state="normal", style="TEntry")
    sezione_entry.configure(state="normal", style="TEntry" )
    data_entry.configure(state="normal", style="TEntry")
    domande_entry.configure(state="normal", style="TEntry")
    intestazione_testo_label.configure(style="TLabel")
    cdl_label.configure(style="TLabel")
    anno_label.configure(style="TLabel")
    sezione_label.configure(style="TLabel")
    data_label.configure(style="TLabel")
    domande_label.configure(style="TLabel")
    button.configure(style="TButton", state="normal")

def cerca_file():
    root.file = filedialog.askopenfilename(initialdir = percorso,title = "Sfoglia",
                                               filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
    filename = root.file.split("/")[-1]
    archivio.set(filename)
    archivio_scelta.configure(text=filename)
    if archivio.get() != "":
        attiva(None)
    domande_var.set(max_domande(root.file))


def salva_pdf():
    nome_pdf = filedialog.asksaveasfilename(initialdir = percorso, initialfile = "ESAME",title = "Salva con nome", 
                                            filetypes=(("PDF","*.pdf"),("all files","*.*")), defaultextension = ".pdf")
    return nome_pdf
def salva_correttore(pdf):
    correttore = pdf.split("/")
    correttore_1 = correttore[0:len(correttore)-1]
    correttore_1 = f"\\".join(correttore_1)
    correttore_nome = correttore_1 + "\\CORRETTORE.pdf"
    return correttore_nome

def max_domande(a):
    file = pd.read_excel(a)
    domande = file['DOMANDA']
    return len(domande)

def genera_pdf():
    # Leggi il file Excel
    file = pd.read_excel(root.file)

    domande = file['DOMANDA']
    corretta =file['RISPOSTA CORRETTA']
    errata1 = file['Testo2']
    errata2 = file['Testo3']
    errata3 = file['Testo4']
    errata4 = file['Testo5']

    # ARRAY
    risposte = [corretta, errata1, errata2, errata3, errata4]
    lettere = ['a', 'b', 'c', 'd', 'e']

    numero_domande = 100

    # INPUTS
    intestazione_testo = intestazione_testo_var.get().upper()
    cdl = cdl_var.get().upper()
    anno = anno_var.get().upper()
    sezione = sezione_var.get().upper()
    data = data_var.get().upper()
    numero_domande = int(domande_var.get())
    if numero_domande > max_domande(root.file):
        messagebox.showerror("Errore", "Numero di domande superiore al numero di domande nel file excel")
    
    indice_domanda = random.sample(range(len(domande)), numero_domande)


    def stampa():
        
        nome_file_pdf = salva_pdf()
        doc = SimpleDocTemplate(nome_file_pdf, pagesize=A4, rightMargin=18, leftMargin=18, topMargin=36, bottomMargin=36)
        nome_correttore_pdf = salva_correttore(nome_file_pdf)
        doc_corretto = SimpleDocTemplate(nome_correttore_pdf, pagesize=A4, rightMargin=18, leftMargin=18, topMargin=36, bottomMargin=36)
        numero_domanda = 0

        # Lista dei paragrafi
        paragraphs = []
        paragraphs_correttore = []
        spazio = Paragraph("-", spazio_style)

        # Intestazione
        intestazione = [
            Paragraph("Nome e cognome: ", bold_style),
            Paragraph("Matricola: ", bold_style),
            Paragraph(f"ESAME DI {intestazione_testo}", centered_style),
            spazio,
            Paragraph(f"CDL di {cdl} - ANNO {anno}", sub_centered_style),
            Paragraph(f"SEZIONE {sezione}", sub_centered_style),
            Paragraph(f"APPELLO DEL {data}", sub_centered_style),
            spazio
        ]
        paragraphs.extend(intestazione)
        paragraphs_correttore.extend(intestazione)

        try:
            # Aggiungi le domande e le risposte al PDF
            for indice in indice_domanda:
                numero_domanda += 1
                domanda_text = f"{numero_domanda}. {domande[indice]}"
                domanda_paragraph = Paragraph(domanda_text, bold_style)
                paragraphs.append(KeepTogether([domanda_paragraph]))
                paragraphs_correttore.append(KeepTogether([domanda_paragraph]))
                indici_risposte = random.sample(range(5), 5)  # Genera 4 indici unici casuali tra 0 e 3
                for i, lettera in enumerate(lettere):
                    indice_risposta = indici_risposte[i]
                    risposta_text = f"[ ] {lettera}. {risposte[indice_risposta][indice]}"
                    risposta_paragraph = Paragraph(risposta_text, normal_style)
                    paragraphs.append(KeepTogether([risposta_paragraph]))
                    if indice_risposta == 0:
                        risposta_corretta_paragraph = Paragraph(risposta_text, corretta_style)
                        paragraphs_correttore.append(KeepTogether([risposta_corretta_paragraph]))
                        continue
                    paragraphs_correttore.append(KeepTogether([risposta_paragraph]))
                paragraphs.append(spazio)
                paragraphs_correttore.append(spazio)

            # Aggiungi i paragrafi al PDF
            doc_corretto.build(paragraphs_correttore)
            doc.build(paragraphs)
            messagebox.showinfo("Domande", "PDF generato con successo!")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante la generazione del PDF: {str(e)}")
    stampa()
    
# ==================VARIABILI================
archivio = StringVar()
intestazione_testo_var = StringVar()
cdl_var = StringVar()
anno_var = StringVar()
sezione_var = StringVar()
data_var = StringVar()
domande_var = StringVar()
# ================ARCHIVIO================
archivio_scelta = ttk.Label(root, text="Scegli l'archivio", background=palette[6])

# =================LABELS=================
intestazione_testo_label = ttk.Label(root, text="MATERIA", style="Off.TLabel")
cdl_label = ttk.Label(root, text="CDL", style="Off.TLabel")
anno_label = ttk.Label(root, text="ANNO", style="Off.TLabel")
sezione_label = ttk.Label(root, text="SEZIONE", style="Off.TLabel")
data_label = ttk.Label(root, text="DATA", style="Off.TLabel")
domande_label = ttk.Label(root, text="DOMANDE", style="Off.TLabel")
# ==================ENTRY=================
intestazione_testo_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=intestazione_testo_var)
cdl_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=cdl_var)
anno_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=anno_var)
sezione_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=sezione_var)
data_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=data_var)
domande_entry = ttk.Entry(root, state="disabled", style="Off.TEntry", font="Montserrat 13", textvariable=domande_var)

# ==================BUTTONS=================
button = ttk.Button(root, text="Genera PDF", command=genera_pdf, state="disabled", style="Off.TButton")
button.bind("<Enter>", on_enter)
button.bind("<Leave>", on_leave)
button2 = ttk.Button(root, image=icon, command=cerca_file, width=15)
button2.bind("<Enter>", on_enter2)
button2.bind("<Leave>", on_leave2)



# ==================GRID=================
archivio_scelta.grid(row=0, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
intestazione_testo_label.grid(row=1, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")
cdl_label.grid(row=2, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")
anno_label.grid(row=3, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")
sezione_label.grid(row=4, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")
data_label.grid(row=5, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")
domande_label.grid(row=6, column=0, padx=3, pady=2, ipadx=5, ipady=5, sticky="ns")

intestazione_testo_entry.grid(row=1, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
cdl_entry.grid(row=2, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
anno_entry.grid(row=3, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
sezione_entry.grid(row=4, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
data_entry.grid(row=5, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")
domande_entry.grid(row=6, column=1, padx=30, pady=2, ipadx=5, ipady=5, sticky="ew")

button.grid(row=7, column=0, columnspan=3, pady=25)
button2.grid(row=0, column=0, padx=10)




if __name__ == "__main__":
    root.mainloop()
