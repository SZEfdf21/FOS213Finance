import pandas as pd
import openpyxl as opx
import tkinter as tk
from tkinter import ttk


def create_code_selection_window(bedrag, mededeling, bankomschrijving, codes):
    root = tk.Tk()
    root.title("Onbekende verrichting")
    #root.geometry("800x200")

    selected_code = tk.StringVar()

    def on_ok():
        selected_code.set(clicked.get())
        root.quit()

    clicked = tk.StringVar()
    clicked.set(codes[0])

    ttk.Label(root, text="Selecteer een correcte code voor volgende verrichting:").pack(pady=5)
    ttk.Label(root, text=f"€{bedrag}; mededeling: {mededeling}").pack()
    ttk.Label(root, text=f"omschrijving: {bankomschrijving}").pack()
    task_menu = ttk.Combobox(root, textvariable=clicked, values=codes)
    task_menu.pack()

    tk.Button(root, text="OK", command=on_ok).pack(pady=10)

    # Ensure the window appears on top
    root.lift()
    root.attributes('-topmost', True)
    root.after_idle(root.attributes, '-topmost', False)
    root.mainloop()
    root.destroy()

    return selected_code.get()


def goThroughExport(exportDf, codes, omschrijvingen):
    jaarrekeningDf = pd.DataFrame(columns=['Code', 'Omschrijving', 'Datum', 'In',
                                           'Uit', 'Saldo', 'Tegenpartij', 'Mededeling'])
    for index, row in exportDf.iterrows():
        # Het type van het bedrag correct maken voor later in de excel   MOET NAAR 1 KOLOM "BEDRAG" VERANDERD WORDEN
        if row['Credit'] != 0:
            row['Credit'] = float(str(row['Credit']).replace(",", "."))
        if row['Debet'] != 0:
            row['Debet'] = float(str(row['Debet']).replace(",", "."))
        # soort overschrijving checken
        typeOverschrijving = checkType(row['Omschrijving'], str(row['vrije mededeling']), row['Bedrag'], codes, omschrijvingen)
        # gegevens toevoegen aan dataframe voor later in de jaarrekening te steken
        jaarrekeningDf.loc[index] = [typeOverschrijving[0], typeOverschrijving[1], row['Datum'], row['Credit'],
                                     row['Debet'], float(row['Saldo'].strip().replace(',', '.')),
                                     row['Naam tegenpartij'], row['vrije mededeling']]
    return jaarrekeningDf


def checkType(bankomschrijving, mededeling, bedrag, codes, omschrijvingen):
    try:
        code = int(mededeling[:3])
    except ValueError:
        code = 0
    # Kijken of de codes in de lijst voorkomen of niet:
    if code in codes:
        omschrijving = omschrijvingen[codes.index(code)]
    # Herkennen van bankkosten:
    elif (bedrag == '-3,86' or bedrag == '-1,21')and code == 0:
        omschrijving = 'Bankkosten'
        code = 700
    else:
        # Codes met omschrijvingen maken voor in de terminal
        codesMetOmschrijving = [f"{code} {omschrijving}" for code, omschrijving in zip(codes, omschrijvingen)]

        selected_code = create_code_selection_window(bedrag, mededeling, bankomschrijving, codesMetOmschrijving)

        # Kijken of er een code is gevonden via de terminal
        try:
            code = int(selected_code[:3])
        except (TypeError, ValueError):
            code = 0

        if code == 0:
            omschrijving = 'Geen overeenkomst'
        else:
            omschrijving = omschrijvingen[codes.index(int(selected_code.split()[0]))]
    return int(code), omschrijving


def appendToJaarrekening(jaarrekening, jaarrekeningDf):
    jaarrekeningLijst = jaarrekeningDf.values.tolist()
    wb = opx.load_workbook(jaarrekening)
    ws = wb["Jaarrekening"]
    for row in jaarrekeningLijst:
        ws.append(row)
    wb.save(filename=jaarrekening)
    return


if __name__ == '__main__':
    jaarrekening = "Excels/2324Jaarrekening.xlsx"
    exportBank = "Excels/BE13734022778639_14-04-2024_tot_13-05-2024.csv"

    # Dataframe maken van de export csv
    exportDf = pd.read_csv(exportBank, sep=';', index_col=False, encoding='latin-1')
    # NaN waarden in dataframe naar lege string
    exportDf.fillna(0, inplace=True)

    # Gebruikte codes uit de jaarrekening excel ophalen
    codeDf = pd.read_excel(jaarrekening, index_col=None, header=None, sheet_name="OverzichtCodes")
    # Lijst met mogelijke codes en omschrijvingen maken:
    codes = codeDf[0].tolist()
    omschrijvingen = codeDf[1].tolist()

    appendToJaarrekening(jaarrekening, goThroughExport(exportDf, codes, omschrijvingen))
