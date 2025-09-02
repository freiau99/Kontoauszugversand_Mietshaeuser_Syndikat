from flask import Flask, render_template, request, redirect, url_for, session, Response
from werkzeug.utils import secure_filename
import os, pickle, datetime
import pandas as pd
import csv, days360
from reportlab.pdfgen.canvas import Canvas
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

app = Flask(__name__)
app.secret_key = "supersecretkey"

UPLOADS_FOLDER = "uploads"
OUTPUT_FOLDER = "Kontoauszüge"
os.makedirs(UPLOADS_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def dkv2_daten_vorbereiten(file, yr):
    """
    Get the dkv2-direktkreditgeberinnen-table into the correct format that can be used for the other functions available in the package.
    
    Necessary libraries: csv, datetime, days360, pandas
    
    params:file CSV-file that contains the information on the Direktkreditgeber:Innen/ subordinated loan givers. File has to be exported from DKVII Direktkreditverwaltung as csv and not changed since then.
    params:yr Year for  which the documents need to be sent.
    """

    #Tabelle einlesen:
    data = open(file, encoding="utf-8")
    tabelle = csv.reader(data, delimiter=";")

    #Tabelleninformationen formatieren:
    data_zeilen = list(tabelle)
    for zeile in data_zeilen[1:]:
        if len(zeile[5]) < 5:
            zeile[5] = "0" + zeile[5] #4stelligen PLZ eine fünfte Stelle hinzufügen
        for eintrag in range(0, len(zeile)):
            if zeile[eintrag][0] == " ":
                zeile[eintrag] = zeile[eintrag][1:] #Leerzeichen zu Beginn einer Zelle löschen
        zeile[11] = zeile[11].replace('\xa0', ' ') #Codierungsfehler löschen
        zeile[12] = datetime.datetime.strptime(zeile[12], "%d.%m.%Y") #Zum Datum umtransformieren.
        zeile[12] = zeile[12].date()
        if zeile[12].year < yr:
            zeile[12] = datetime.date(yr, 1, 1) #Wenn die Einzahlung in der Vergangenheit liegt, wird das Anfangsdatum des aktuellen Jahres verwendet.
        zeile[14] = datetime.datetime.strptime(zeile[14], "%d.%m.%Y")
        zeile[14] = zeile[14].date()
        if zeile[14].year > 2025:
            zeile[14] = datetime.date(2025, 12, 31) #Wenn unbefristet Enddatum des aktuellen Jahres.
        zeile.append(days360.days360(zeile[12], zeile[14], method="EU")) #TAGE360 berechnen - TAGE360 errechnet die Anzahl der zu verzinsten Tage!
        zeile.append(zeile[11].replace(' €', '')) #Eurozeichen entfernen - notwendig für das Zahlenformat
        zeile[-1] = zeile[-1].replace('.', '') #Tausenderpunkte entfernen - notwendig für das Zahlenformat (In Python und weiteren Statistikprogrammen ist immer ein . der Dezimaltrenner!)
        zeile[-1] = zeile[-1].replace(',', '.') #Kommas durch Punkte ersetzen - In Python ist ein . der Dezimalstellentrenner!
        zeile[-1] = float(zeile[-1])
        zeile.append(float(zeile[10][:4].replace(",", ".")))
        zeile.append(str(round((zeile[-2] + zeile[-2] * (zeile[-1] / 100) * (zeile[-3] / 359)), 2))  + " €") #Endbetrag berechnen
        #Hiermit habe ich die Tabelle so weit fertig, wie es vorerst nötig ist. 
    data.close() 

    direktis = {f"{zeile[2]} {zeile[3]}": 
                list([zeile[4], zeile[5] + " " + zeile[6], zeile[7], zeile[11], zeile[10], zeile[15], zeile[-1]]) 
                for zeile in data_zeilen[1:]} #Dieses Objekt ist wichtig für die spätere Erstellung der Kontoauszüge - die nötigen Infos werden hieraus gezogen.
    
    # Umwandlung in DataFrame
    direktis_tabelle = pd.DataFrame.from_dict(
        direktis, 
        orient='index', 
        columns=["Haus", "PLZ und Ort", "Mailadresse", "Betrag", "Zinssatz", "Zinsmethode", "Neuer Betrag"]
    )
    direktis_tabelle.index.name = "Name"
    
    return direktis, direktis_tabelle

def kontoauszüge_erstellen(yr, projektname, projektartikel, direktis):
    """
    This function creates unique subordinated loan giver-specific PDF-format documents to give the information necessary for the Kontoauszüge/ bank statements.
    
    Necessary libraries: reportlab
    
    params:yr Year of the Kontoauszug/ bank statement
    params:projektname Name of the submitting Syndikats-project.
    params:projektartikel Article of the submitting Syndikats-project. "Eure", "Die", "Das Haus" etc.
    params:direktis direktis-object resulting from the correct use of the dkv2_daten_vorbereiten-function.
    """
    
    
    for geber, values in direktis.items():
        filename = os.path.join(OUTPUT_FOLDER, f"Kontoauszug_{geber.replace(' ','_')}.pdf")
        c = Canvas(filename)
        textobject = c.beginText(36, 800)
        lines = [
            f"Hallo {geber}",
            "",
            "Vielen Dank für Deine Unterstützung!",
            f"Hiermit übersenden wir Dir Deinen Kontoauszug des Jahres {yr}:",
            "",
            f"Name: {geber}",
            f"Jahr: {yr}",
            f"Betrag Anfang des Jahres: {values[3]}",
            f"gewählter Zinssatz: {values[4]}",
            f"Zinsmethode: {values[5]}",
            f"Betrag zum Ende des Jahres: {values[6]}",
            "",
            "Bei Fragen, Anmerkungen und Ergänzungen antworte uns gerne.",
            "",
            "Fröhliche Festtage und einen guten Rutsch",
            "",
            f"{projektartikel} {projektname}"
        ]
        for line in lines:
            textobject.textLine(line)
        c.drawText(textobject)
        c.showPage()
        c.save()


def kontoauszüge_versenden(direktis, year, projektname, projektartikel,
                            smtp_server, email_adresse, passwort,
                            betreff, mailtext_template, directory, prefix="kontoauszug_"):
    """
    Function to send the bank statements to the addressed persons. 
    
    Required libraries smtplib, getpass, email, IPython.
    
    params:direktis list object containing the mail accounts of the addressed persons.
    params:year Year for which the bank statements should be done.
    params:projektname name of the project which sends the bank statements - string.
    params:projektartikel article of the project which sends the bank statements, i.e., "Eure", "Die"
    params:smtp_server smtp server of the user's e-mail address.
    params:email_adresse user e-mail address.
    params:passwort password used to login in the user's e-mail account.
    params:betreff subject of the mail.
    params:mailtext_template template used for the mail's text body
    params:directory directory in which the bank statements are stored
    params:prefix prefix of the bank statements, defaults to "kontoauszug_"
    """
    versand_status = {}
    for geber, values in direktis.items():
        try:
            body = mailtext_template.format(
                name=geber, projektname=projektname, projektartikel=projektartikel, yr=year
            )
            message = MIMEMultipart()
            message['Subject'] = betreff
            message['From'] = email_adresse
            message['To'] = values[2]
            message.attach(MIMEText(body))

            filename = f"{directory}/{prefix}{geber.replace(' ','_')}.pdf"
            with open(filename, 'rb') as file:
                message.attach(MIMEApplication(file.read(), Name=os.path.basename(filename)))

            smtp_objekt = smtplib.SMTP(smtp_server, 587)
            smtp_objekt.starttls()
            smtp_objekt.login(email_adresse, passwort)
            smtp_objekt.sendmail(email_adresse, values[2], message.as_string())
            smtp_objekt.quit()

            versand_status[geber] = "Kontoauszug versendet"
        except Exception as e:
            versand_status[geber] = f"Kontoauszug nicht versendet: {str(e)}"
    return versand_status
    


# --- Flask-Routen ---

@app.route("/", methods=["GET", "POST"])
def home():
    if "step" not in session:
        session["step"] = 1

    if request.method == "POST":
        if "back" in request.form:
            session["step"] = max(1, session["step"] - 1)
            return redirect(url_for("home"))

        # STEP 1: Upload CSV
        if session["step"] == 1 and "next" in request.form:
            csv_file = request.files.get("csv_file")
            year = request.form.get("year", type=int)   # <<< NEU (sicher casten)
            if csv_file and year:
                filepath = os.path.join(UPLOADS_FOLDER, secure_filename(csv_file.filename))
                csv_file.save(filepath)
                session["csv_path"] = filepath          # <<< NEU (nur Pfad speichern)
                session["year"] = year                  # <<< NEU
                session["step"] = 2                     # <<< NEU
                return redirect(url_for("home"))

        # STEP 2: Projektangaben
        if "next" in request.form and session["step"] == 2:
            session["projektname"] = request.form["projektname"]
            session["projektartikel"] = request.form["projektartikel"]

            # <<< NEU: direktis bei Bedarf frisch berechnen (nicht in Session speichern)
            direktis, _ = dkv2_daten_vorbereiten(session["csv_path"], int(session["year"]))

            kontoauszüge_erstellen(session["year"], session["projektname"], session["projektartikel"], direktis)
            session["step"] = 3                         # <<< GEÄNDERT: explizit auf 3 setzen
            return redirect(url_for("home"))

        # STEP 3: Mailversand
        if "next" in request.form and session["step"] == 3:
            session["smtp_server"] = request.form["smtp_server"]
            session["email_adresse"] = request.form["email_adresse"]
            session["passwort"] = request.form["passwort"]
            session["betreff"] = request.form["betreff"]
            session["mailtext_template"] = request.form["mailtext"]

            # <<< NEU: direktis wieder frisch berechnen
            direktis, _ = dkv2_daten_vorbereiten(session["csv_path"], int(session["year"]))

            status = kontoauszüge_versenden(
                direktis, session["year"], session["projektname"], session["projektartikel"],
                session["smtp_server"], session["email_adresse"], session["passwort"],
                session["betreff"], session["mailtext_template"], OUTPUT_FOLDER,
                prefix="Kontoauszug_"                    # <<< WICHTIG: passt zu erzeugten PDF-Namen
            )
            session["versand_status"] = status
            session["step"] = 4                          # <<< GEÄNDERT: explizit auf 4 setzen
            return redirect(url_for("home"))

    # **Immer rendern, egal ob GET oder POST**
    # <<< NEU: Tabelle nur für die Anzeige berechnen (nicht in Session legen)
    table_html = None
    if session.get("step", 1) >= 2 and session.get("csv_path") and session.get("year"):
        try:
            _, tabelle = dkv2_daten_vorbereiten(session["csv_path"], int(session["year"]))
            table_html = tabelle.to_html(classes="table table-striped", border=0)
        except Exception as e:
            print("Warnung Tabellenanzeige:", e)

    return render_template("home.html",
                           step=session["step"],
                           versand_status=session.get("versand_status"),
                           table=table_html)               # <<< GEÄNDERT: aus lokaler Variable

@app.route("/download_status")
def download_status():
    versand_status = session.get("versand_status", {})
    df = pd.DataFrame(list(versand_status.items()), columns=["Name", "Status"])
    csv_data = df.to_csv(index=False, sep=";")
    return Response(csv_data, mimetype="text/csv",
                    headers={"Content-Disposition": "attachment;filename=versand_status.csv"})


if __name__ == "__main__":
    app.run(debug=False)
    