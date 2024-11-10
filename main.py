import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import jaarrekeningUpdater as jU
import checkBetaling as cB


# Functie om bestandslocaties op te halen met Tkinter
def get_file_locations():
    root = tk.Tk()
    root.withdraw()  # Verberg het hoofdvenster

    messagebox.showinfo("Bestanden verzamelen", "Geef de locatie van de volgende bestanden en klik op OK")

    initial_dir = "./Excels"
    jaarrekening = filedialog.askopenfilename(title="Selecteer Jaarrekening", filetypes=[("Excel files", "*.xlsx")], initialdir=initial_dir)
    schuldenoverzicht = filedialog.askopenfilename(title="Selecteer Schuldenoverzicht", filetypes=[("Excel files", "*.xlsx")], initialdir=initial_dir)
    exportBank = filedialog.askopenfilename(title="Selecteer Recentste export bank", filetypes=[("CSV files", "*.csv")], initialdir=initial_dir)

    if not jaarrekening or not schuldenoverzicht or not exportBank:
        messagebox.showerror("Fout", "Alle bestanden moeten geselecteerd worden.")
        root.destroy()
        return None, None, None

    root.destroy()
    return jaarrekening, schuldenoverzicht, exportBank


# Dataframe maken van de export csv
def create_export_df(exportBank):
    exportDf = pd.read_csv(exportBank, sep=';', index_col=False, encoding='latin-1')
    exportDf.fillna(0, inplace=True)
    return exportDf

# Gebruikte codes uit de jaarrekening excel ophalen
def get_codes(jaarrekening):
    codeDf = pd.read_excel(jaarrekening, index_col=None, header=None, sheet_name="OverzichtCodes")
    codes = codeDf[0].tolist()
    omschrijvingen = codeDf[1].tolist()
    return codes, omschrijvingen


def welke_taak():
    root = tk.Tk()
    root.title("Welke Taak?")
    root.geometry("450x150")

    def on_ok():
        selected_task.set(clicked.get())
        root.quit()

    tasks = [
        '0. Niets',
        '1. Jaarrekening updaten',
        '2. Lidgelden, EHWE of Kamp inschrijvingen updaten',
        '3. SchuldenOverzicht updaten'
    ]

    clicked = tk.StringVar()
    clicked.set(tasks[0])
    selected_task = tk.StringVar()

    ttk.Label(root, text="Welke actie wil je uitvoeren?").pack()
    task_menu = ttk.Combobox(root, textvariable=clicked, values=tasks)
    task_menu.pack()

    tk.Button(root, text="Ok", command=on_ok).pack()

    # Venster naar voor laten komen
    root.lift()
    root.attributes('-topmost', True)
    root.after_idle(root.attributes, '-topmost', False)
    root.mainloop()
    root.destroy()
    return int(selected_task.get()[0]) if selected_task.get() else 0


def welke_subtaak():
    def on_ok():
        subtasks["Lidgelden"] = lidgelden_var.get()
        subtasks["EHWE inschrijvingen"] = ehwe_var.get()
        subtasks["Kamp inschrijvingen"] = kamp_var.get()
        root.quit()

    root = tk.Tk()
    root.title("Welke controles wil je uitvoeren?")
    root.geometry("450x150")

    subtasks = {}

    lidgelden_var = tk.BooleanVar()
    ehwe_var = tk.BooleanVar()
    kamp_var = tk.BooleanVar()

    ttk.Checkbutton(root, text="Lidgelden", variable=lidgelden_var).pack(pady=5)
    ttk.Checkbutton(root, text="EHWE inschrijvingen", variable=ehwe_var).pack(pady=5)
    ttk.Checkbutton(root, text="Kamp inschrijvingen", variable=kamp_var).pack(pady=5)

    ttk.Button(root, text="OK", command=on_ok).pack(pady=10)

    root.mainloop()
    root.destroy()
    return subtasks


jaarrekening, schuldenoverzicht, exportBank = get_file_locations()
if jaarrekening and schuldenoverzicht and exportBank:
    exportDf = create_export_df(exportBank)
    codes, omschrijvingen = get_codes(jaarrekening)

TakenGedaan = False
while not TakenGedaan:
    taak = welke_taak()
    if taak == 0:
        TakenGedaan = True
    elif taak == 1:
        jU.appendToJaarrekening(jaarrekening, jU.goThroughExport(exportDf, codes, omschrijvingen))
    elif taak == 2:
        sub_taak = welke_subtaak()
        if sub_taak["Lidgelden"]:
            cB.lidgeldInschrijvingUpdater(jaarrekening, "LidgeldInschrijvingen", 100, 'Lidgeld', 'F')
        if sub_taak["EHWE inschrijvingen"]:
            cB.lidgeldInschrijvingUpdater(jaarrekening, "LidgeldInschrijvingen", 505, 'EHWE', 'H')
        if sub_taak["Kamp inschrijvingen"]:
            cB.lidgeldInschrijvingUpdater(jaarrekening, "LidgeldInschrijvingen", 605, 'Kamp', 'J')
    elif taak == 3:
        cB.schuldenUpdater(jaarrekening, schuldenoverzicht, "TeBetalen", 300)

print('Goed gedaan de gevraagde taken zijn uitgevoerd!')
