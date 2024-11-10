import pandas as pd
import smtplib
import ssl
from email.message import EmailMessage
import time

# Gegevens van de leiding importeren
excelFile = "Excels/2324Jaarrekening.xlsx"

# Email onderwerp opstellen
# subject_email = 'Herinnering betaling lidgeld 2022-2023'
# subject_email = 'Herinnering betaling eenheidsweekend 2023'
subject_email = 'Herinnering betaling kampinschrijving 2024'

# Mail lijst selecteren testMail of eMail:
mailingList = 'eMail'

# Genereren van dataframe met schulden uit excel:
df = pd.read_excel(excelFile, index_col=None, header=0, sheet_name="LidgeldInschrijvingen")

# Wat bekijken we (Lidgeld, EHWE, Kamp)
wat = 'Kamp'

# rijen weghalen die niet nodig zijn:
df = df[df[(wat + ' Betaald')] == 0]
# df.drop(df.head(46).index,inplace=True)
# df.drop(df.tail(1).index,inplace=True)
# df = df[df['Tak'] == 'Stam']

# De html mail importeren
# htmlMail = 'htmlMails/mailEHWEInschrijving.html'
htmlMail = 'htmlMails/mailKamp.html'
with open(htmlMail, 'r', ) as f:
    tekst = f.read()

# Gegevens zender invullen + mailserver inrichten
port = 587
smtp_server = 'smtp.gmail.com'
sender_name = 'Scouts De Toekan'
login_email = 'nandoe@detoekan.be'
sender_email = 'info@detoekan.be'
password = input("Type your password and press enter: ")
# Outlook: 'sander.devleeschauwer@outlook.com', port: 587, smtp_server = 'smtp.office365.com'
# Gmail: 'sanderdevleeschauwer318@gmail.com', port: 587, smtp_server = 'smtp.gmail.com'
# Gmail scouts: 'nandoe@detoekan.be', 'penningmeester@detoekan.be', port: 587, smtp_server = 'smtp.gmail.com'

# Mail opstellen per rij in de dataframe:
start = time.time()
aantalMails = len(df.index)
counter = 0
for index, row in df.iterrows():
    # Aanpassen naar wat er in de html file nodig is (welke schulden enzo):
    receiver_email = row[mailingList]
    voornaam = row['Voornaam']
    achternaam = row['Achternaam']
    bedrag = str(row[wat]).replace('.', ',')

    aangepasteTekst = tekst.format(voornaam=voornaam, achternaam=achternaam, bedrag=bedrag)

    msg = EmailMessage()
    msg['Subject'] = subject_email
    msg['From'] = sender_name + '<' + sender_email + '>'
    msg['To'] = receiver_email
    msg.set_content(aangepasteTekst, subtype='html')

    context = ssl.create_default_context()

    with smtplib.SMTP(smtp_server, port) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(login_email, password)
        server.send_message(msg)
    counter += 1
    print(counter, '/', aantalMails)

server.close()

end = time.time()
print(end-start)
