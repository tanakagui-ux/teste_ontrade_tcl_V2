from flask import Flask, render_template, request
import pandas as pd
import uuid, os, base64
from datetime import datetime
import threading

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

import smtplib
from email.message import EmailMessage

app = Flask(__name__)
EXCEL_FILE = "dados.xlsx"

# Garante que o Excel existe
if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=[
        "Token","Data","Promotor","Supervisor","Email","Status","Assinado"
    ]).to_excel(EXCEL_FILE, index=False)


@app.route('/')
def form():
    return render_template("form.html")


@app.route('/submit', methods=['POST'])
def submit():
    token = str(uuid.uuid4())

    data = {
        "Token": token,
        "Data": request.form['data'],
        "Promotor": request.form['promotor'],
        "Supervisor": request.form['supervisor'],
        "Email": request.form['email'],
        "Status": request.form['status'],
        "Assinado": "Pendente"
    }

    # Salvar no Excel
    df = pd.read_excel(EXCEL_FILE)
    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    link = request.host_url + "sign/" + token

    # 🔥 ENVIO DE EMAIL EM BACKGROUND (CORRIGIDO)
    threading.Thread(
        target=enviar_email_link,
        args=(data["Email"], link)
    ).start()

    return "Solicitação enviada para assinatura!"


@app.route('/sign/<token>')
def sign(token):
    return render_template("sign.html", token=token)


@app.route('/save-signature', methods=['POST'])
def save_signature():
    data = request.json
    token = data['token']
    image = data['image']

    os.makedirs("assinaturas", exist_ok=True)
    file_path = f"assinaturas/{token}.png"

    image_data = base64.b64decode(image.split(',')[1])

    with open(file_path, "wb") as f:
        f.write(image_data)

    df = pd.read_excel(EXCEL_FILE)
    df.loc[df["Token"] == token, "Assinado"] = "Sim"
    df.to_excel(EXCEL_FILE, index=False)

    gerar_pdf(token)

    return "OK"


def gerar_pdf(token):
    df = pd.read_excel(EXCEL_FILE)
    row = df[df["Token"] == token].iloc[0]

    file_name = f"termo_{row['Promotor'].replace(' ', '_')}.pdf"

    doc = SimpleDocTemplate(
        file_name,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=3*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    elements = []

    # Logo
    if os.path.exists("Quadrado_fundo_branco.png"):
        logo = Image("Quadrado_fundo_branco.png", width=3*cm, height=3*cm)
        logo.hAlign = 'RIGHT'
        elements.append(logo)

    elements.append(Spacer(1, 20))

    texto = f"""
    Atestamos para os devidos fins que o(a) promotor(a) {row['Promotor']} participou do processo de Capacitação Inicial Ontrade | TCL.

    O participante demonstrou capacidade e está classificado como: <b>{row['Status'].upper()}</b>.

    <br/><br/>
    Responsável: {row['Supervisor']}<br/>
    Empresa: Ontrade | TCL<br/>
    Data: {row['Data']}
    """

    elements.append(Paragraph(texto, styles["Normal"]))

    # Assinatura
    assinatura_path = f"assinaturas/{token}.png"
    if os.path.exists(assinatura_path):
        elements.append(Spacer(1, 30))
        elements.append(Image(assinatura_path, width=6*cm, height=3*cm))

    doc.build(elements)

    # Enviar email com PDF
    threading.Thread(target=enviar_email, args=(file_name,)).start()


def enviar_email_link(destinatario, link):
    try:
        msg = EmailMessage()
        msg['Subject'] = 'Assinatura pendente'
        msg['From'] = os.environ.get("EMAIL_USER")
        msg['To'] = destinatario
        msg.set_content(f"Clique para assinar:\n{link}")

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get("EMAIL_USER"), os.environ.get("EMAIL_PASS"))
            smtp.send_message(msg)
    except Exception as e:
        print("Erro ao enviar email:", e)


def enviar_email(arquivo):
    try:
        msg = EmailMessage()
        msg['Subject'] = 'PDF Assinado'
        msg['From'] = os.environ.get("EMAIL_USER")
        msg['To'] = "guilherme.tanaka@ontrade.com.br"
        msg.set_content("Segue o PDF assinado.")

        with open(arquivo, 'rb') as f:
            msg.add_attachment(
                f.read(),
                maintype='application',
                subtype='pdf',
                filename=arquivo
            )

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(os.environ.get("EMAIL_USER"), os.environ.get("EMAIL_PASS"))
            smtp.send_message(msg)
    except Exception as e:
        print("Erro ao enviar PDF:", e)


if __name__ == '__main__':
    app.run()
