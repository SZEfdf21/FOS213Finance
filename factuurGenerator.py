import pandas as pd
import numpy as np
from datetime import date
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os


def generateOverzicht(naam, df, facturatieTekst, onderwerp):

    # Juiste rij in Dataframe verkrijgen
    df = df.loc[df['Naam'] == naam]
    # Enkel belangrijke kolommen voor iteratie
    detailDf = df.drop(columns=df.columns[:4], axis=1)

    if df.empty == False:
        # Wat we nodig hebben:
        columnNumber = df.columns.get_loc(onderwerp + " Totaal")
        subtotaal = np.round(float(df[onderwerp + " Totaal"].iloc[0]), 2)
    else:
        subtotaal = 0.0

    # Controle indien 0 euro (om tijd te winnen):
    if subtotaal != 0.0:
        facturatieLijn = "- {uitgave}: &euro; {bedrag} <br>"
        overzicht = ""

        # Detail opstellen
        for kolom in detailDf:
            if "Unnamed" in kolom:
                break
            bedrag = np.round(float(detailDf[kolom].iloc[0]), 2)
            bedrag = str(np.round(bedrag, 2)).replace(".", ",")
            overzicht += facturatieLijn.format(uitgave=kolom, bedrag=bedrag)

        # FacturatieTekst opstellen
        subtotaalTekst = str(subtotaal).replace(".", ",")
        facturatieTekstIngevuld = facturatieTekst.format(onderwerp=onderwerp, subtotaal=subtotaalTekst, overzicht=overzicht)
    else:
        subtotaal = 0.0
        facturatieTekstIngevuld = ""

    return [facturatieTekstIngevuld, subtotaal]


def generateFactuur(naam, mailAdres, titel, due_date, reedsBetaald, facturatieTekst, poefDf, lwDf, inleefweekDf, ehweDf,
                    kampDf, vaDf, vjDf):
    # Overzichten aanmaken:
    poefOverzicht = generateOverzicht(naam, poefDf, facturatieTekst, "Poeflijsten")
    lwOverzicht = generateOverzicht(naam, lwDf, facturatieTekst, "LW")
    inleefweekOverzicht = generateOverzicht(naam, inleefweekDf, facturatieTekst, "Inleefweek")
    ehweOverzicht = generateOverzicht(naam, ehweDf, facturatieTekst, "EHWE")
    kampOverzicht = generateOverzicht(naam, kampDf, facturatieTekst, "Kamp")
    vaOverzicht = generateOverzicht(naam, vaDf, facturatieTekst, "Voorschotten allerlei")
    vjOverzicht = generateOverzicht(naam, vjDf, facturatieTekst, "Vorig jaar")

    # Paden voor afbeeldingen ophalen (werkt niet helemaal):
    #toekan = "file:///" + os.path.abspath("./Facturen/Logo_Kaderblauw.png")
    #qrPath = "file:///" + os.path.abspath("./Facturen/QR_Totale schuld.png.png")

    # Opstellen totaal overzicht:
    facturatieTekstVol = (poefOverzicht[0] + lwOverzicht[0] + inleefweekOverzicht[0] + ehweOverzicht[0]
                          + kampOverzicht[0] + vaOverzicht[0] + vjOverzicht[0])
    somSubtotalen = (poefOverzicht[1] + lwOverzicht[1] + inleefweekOverzicht[1] + ehweOverzicht[1]
                     + kampOverzicht[1] + vaOverzicht[1] + vjOverzicht[1])

    somSubtotalen = np.round(somSubtotalen, 2)
    reedsBetaald = np.round(reedsBetaald, 2)
    totaalSchuld = np.round(somSubtotalen - reedsBetaald, 2)

    # print(facturatieTekstVol)
    # print(somSubtotalen)

    # Alles in html template steken:
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('./Facturen/leeg_invoice.html')

    rendered_html = template.render(naam=naam, mailAdres=mailAdres, titel=titel, date=date.today().strftime("%d/%m/%Y"),
                                    due_date=due_date, reedsBetaald=str(reedsBetaald).replace(".",","),
                                    facturatieTekstVol=facturatieTekstVol,
                                    totaalSchuld=str(totaalSchuld).replace(".",","))

    output_path = os.path.join("./Facturen/", f"Factuur.html")
    with open(output_path, 'w') as file:
        file.write(rendered_html)

    # Template omzetten naar PDF (Correcte installatie Weasyprint)
    #   Installeer https://www.gtk.org/docs/installations/windows#using-gtk-from-msys2-packages
    #   Na installatie volgende runnen:
    #       "pacman -S mingw-w64-x86_64-gtk3 mingw-w64-x86_64-cairo mingw-w64-x86_64-gobject-introspection"
    #   Bij omgevingsvariabelen ook volgende zetten in path: "C:\msys64\mingw64\bin" en "C:\msys64\ucrt64\bin"
    #   Computer herstarten
    html = HTML(f'./Facturen/Factuur.html')
    html.write_pdf(f'./Facturen/Factuur.pdf')
    return


if __name__ == '__main__':
    # Leeg stukje template voor factuur overzicht
    htmlTekst = 'Facturen/leeg_overzicht.html'
    with open(htmlTekst, 'r', ) as f:
        facturatieTekst = f.read()

    # Dataframe's met info verkrijgen:
    schuldenOverzicht = "Excels/2425SchuldenOverzicht.xlsx"
    # Poeflijsten df:
    poefDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="Poeflijsten")
    poefDf.drop(columns=poefDf.columns[-10:], axis=1, inplace=True)
    poefDf.fillna(0, inplace=True)
    # LW df:
    lwDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="LW")
    lwDf.fillna(0, inplace=True)
    # Inleefweek df:
    inleefweekDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="Inleefweek")
    inleefweekDf.fillna(0, inplace=True)
    # EHWE df:
    ehweDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="EHWE")
    ehweDf.fillna(0, inplace=True)
    # Kamp df:
    kampDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="Kamp")
    kampDf.fillna(0, inplace=True)
    # Voorschotten allerlei df:
    vaDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="Voorschotten allerlei")
    vaDf.fillna(0, inplace=True)
    # Vorig jaar df:
    vjDf = pd.read_excel(schuldenOverzicht, index_col=None, header=0, sheet_name="Vorig jaar")
    vjDf.fillna(0, inplace=True)

    generateFactuur("Caracara", "caracara@detoekan.be", "Schuldenoverzicht van 345e FOS De Toekan (update: mei)", "22/06/2024",228.81, facturatieTekst, poefDf, lwDf, inleefweekDf, ehweDf, kampDf, vaDf, vjDf)
