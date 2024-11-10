import pandas as pd
import factuurGenerator as fg
from email.message import EmailMessage
import smtplib
import ssl
import time
import xlwings
import qrcode


####################################
###           Inputs             ###
####################################

# Gegevens zender invullen + mailserver inrichten
sender_name = 'Uw Penningmeester'
login_email = 'nandoe@detoekan.be'
sender_email = 'nandoe@detoekan.be'
password = input("Typ uw App-passwoord en druk enter: ")
# Outlook: 'sander.devleeschauwer@outlook.com', port: 587, smtp_server = 'smtp.office365.com'
# Gmail: 'sanderdevleeschauwer318@gmail.com', port: 587, smtp_server = 'smtp.gmail.com'
# Gmail scouts: 'nandoe@detoekan.be', 'penningmeester@detoekan.be', port: 587, smtp_server = 'smtp.gmail.com'

# Gegevens van de leiding importeren
excelFile = "Excels/2324SchuldenOverzicht.xlsx"

# Email onderwerp opstellen
subject_email = 'Schuldenoverzicht bij 345e FOS de Toekan (update: juli)'
due_date = "30/08/2024"

# Opgeven van de leidingInformatie
leidingInfo = f"""Graag betalen voor {due_date}."""

# Mail lijst selecteren (gebruik 'testMail' voor te testen of 'eMail' voor echt):
mailingList = 'eMail'


# Stukje code voor Excel zodanig te fixen dat we dat deftig kunnen lezen zonder formules
excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open(excelFile)
excel_book.save()
excel_book.close()
excel_app.quit()
df = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="TeBetalen")


####################################
###     Selecteren naar wie      ###
####################################

# Mail versturen naar alle leiding en seniors, bij stam enkel naar zij die schulden hebben:
df = df.drop(df[(df['Type'] == 'Stam') & (df['Totale schulden'] == 0)].index)

# Kan handig zijn voor iets weg te halen of 1 iemand te kiezen:
# df.drop(df.head(40).index,inplace=True)
# df.drop(df.tail(1).index,inplace=True)
df.drop(df[df['Type'] == 'Seniors'].index, inplace=True)
df.drop(df[df['Type'] == 'Leiding'].index, inplace=True)
# df.drop(df[df['Naam'] == 'Pekari'].index, inplace=True)
# df = df[(df['Naam'] == 'Coati')]


####################################
###          Aflblijven          ###
####################################

# nan waarden naar float:
df = df.fillna(0)

# Afronden van alle waarden in de df
df = df.round({'Poeflijsten': 2, 'LW': 2, 'Inleefweek': 2, 'Kamp': 2, 'Voorschotten allerlei': 2,
               'Vorig jaar': 2, 'Totale schulden': 2, 'Rekening In': 2, 'Cash In': 2})


# De html mail importeren
htmlMail = 'htmlMails/mailSchuldenOverzicht.html'
with open(htmlMail, 'r', ) as f:
    tekst = f.read()

# De lege_tabel importeren voor later (dit document niet aanpassen!)
htmlMail = 'htmlMails/lege_tabel.html'
with open(htmlMail, 'r', ) as f:
    legeTabel = f.read()

# Leeg stukje template voor factuur overzicht
    htmlTekst = 'Facturen/leeg_overzicht.html'
    with open(htmlTekst, 'r', ) as f:
        facturatieTekst = f.read()

# Poeflijsten df:
poefDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="Poeflijsten")
poefDf.drop(columns=poefDf.columns[-10:], axis=1, inplace=True)
poefDf.fillna(0, inplace=True)
# LW df:
lwDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="LW")
lwDf.fillna(0, inplace=True)
# Inleefweek df:
inleefweekDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="Inleefweek")
inleefweekDf.fillna(0, inplace=True)
# EHWE df:
ehweDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="EHWE")
ehweDf.fillna(0, inplace=True)
# Kamp df:
kampDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="Kamp")
kampDf.fillna(0, inplace=True)
# Voorschotten allerlei df:
vaDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="Voorschotten allerlei")
vaDf.fillna(0, inplace=True)
# Vorig jaar df:
vjDf = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="Vorig jaar")
vjDf.fillna(0, inplace=True)


# QR generator
def qrGenerator(bedrag, mededeling, onderwerp):
    opslagLocatie = "Facturen/QR_" + onderwerp + ".png"
    bedrag = bedrag.replace(',','.')
    codeInfo = "BCD\n001\n1\nSCT\nKREDBEBB\n345E FOS DE TOEKAN\nBE13734022778639\nEUR" + bedrag + "\n\n" + mededeling
    code = qrcode.make(codeInfo)
    code.save(opslagLocatie)
    return


# Mailverzendfunctie
def sendMail(subject_email, sender_name, sender_email, reciever_email, bericht, login_email, password):
    port = 587
    smtp_server = 'smtp.gmail.com'
    msg = EmailMessage()
    msg['Subject'] = subject_email
    msg['From'] = sender_name + '<' + sender_email + '>'
    msg['To'] = reciever_email
    msg.set_content(bericht, subtype='html')

    # PDF-overzicht toevoegen
    FileName = "Facturen/Factuur.pdf"
    with open(FileName, 'rb') as file:
        content = file.read()
        msg.add_attachment(content, maintype='application', subtype='pdf', filename="Schuldenoverzicht.pdf")

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(login_email, password)
        server.send_message(msg)
    server.close()
    return


start = time.time()

# Mail per mail opstellen
aantalMails = len(df.index)
counter = 0
for index, row in df.iterrows():
    # Aanpassen naar wat er in de html file nodig is (welke schulden enzo):
    reciever_email = row[mailingList]
    naam = row['Naam']
    reedsBetaald = row['Rekening In'] + row['Cash In']
    totaleSchuld = str(row['Totale schulden']).replace('.', ',')
    onderwerp = "Totale schuld"
    bedrag = totaleSchuld
    mededeling = "300 {}".format(naam)

    # Tabel genereren
    tabellen = ""
    if row['Totale schulden'] <= 0:
        tabellen = """Proficiat, u heeft geen openstaande schulden bij de scouts! 
        De penningmeester dankt u voor het vertrouwen in onze firma."""

        # QR-code en Factuur genereren
        qrGenerator(bedrag, mededeling, onderwerp)
        fg.generateFactuur(naam, reciever_email, subject_email, due_date, reedsBetaald, facturatieTekst,
                           poefDf, lwDf, inleefweekDf, ehweDf, kampDf, vaDf, vjDf)

    else:
        # Alles toevoegen aan de mail
        tabel = legeTabel.format(bedrag=bedrag, mededeling=mededeling)
        tabellen = tabellen + tabel

        # QR-code en Factuur genereren
        qrGenerator(bedrag, mededeling, onderwerp)
        fg.generateFactuur(naam, reciever_email, subject_email, due_date, reedsBetaald, facturatieTekst,
                           poefDf, lwDf, inleefweekDf, ehweDf, kampDf, vaDf, vjDf)
        

    # Extra info voor leiding toevoegen (bv: tegen wanneer betalen)
    if row['Type'] == 'Leiding':
        leidingInfo = leidingInfo
    else:
        leidingInfo = ""

    # Finale mail opstellen
    aangepasteTekst = tekst.format(naam=naam, leidingInfo=leidingInfo, tabellen=tabellen)

    # Mail verzenden
    sendMail(subject_email, sender_name, sender_email, reciever_email, aangepasteTekst, login_email, password)
    counter += 1
    print(counter, '/', aantalMails)

end = time.time()
print(end-start)
