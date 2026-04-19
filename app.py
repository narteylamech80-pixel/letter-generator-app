from flask import Flask, render_template, request, redirect, url_for, flash
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from email.message import EmailMessage
import smtplib
import pandas as pd
import ssl
import os
from flask import send_file

import smtplib
from email.message import EmailMessage


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.makedirs(os.path.join(BASE_DIR, "letters"), exist_ok=True)




# function to generate letter
def generate_letter(name, student_id, programme, phone):

    print(f"Generating letter for {name}")

    # ======================
# FILE OUTPUT
# ======================


    def draw_red_line(canvas, doc):
        
        canvas.setStrokeColor(colors.red)   # set line color
        canvas.setLineWidth(2)              # line thickness
        
        canvas.line(
            2*cm, 23*cm,   # starting point (x1, y1)
            20*cm, 23*cm   # ending point (x2, y2)
        )
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1.3)
        canvas.line(
            2*cm, 5.25*cm,   # starting point (x1, y1)
            20*cm, 5.25*cm   # ending point (x2, y2)
        )







    font_path = os.path.join(BASE_DIR, "source_files", "fonts", "MAIAN.TTF")
    if "MAIAN" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("MAIAN", font_path))
    
    safe_name = name.replace(" ", "_")
    pdf_path = os.path.join(BASE_DIR, "letters", f"{safe_name}.pdf")

    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=2*cm,
        topMargin=1*cm,
        bottomMargin=2*cm
        
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "title",
        parent=styles["Heading2"],
        alignment=1,
        fontName="MAIAN",
        fontSize=18,
        spaceAfter=10
    )

    title_style_2 = ParagraphStyle(
        "address",
        parent=styles["Heading4"],
        alignment=2,
        fontName="MAIAN",
        fontSize=12,
        spaceAfter=2
    )

    title_style_2_1 = ParagraphStyle(
        "our_reference",
        parent=styles["Heading4"],
        alignment=0,
        fontName="MAIAN",
        fontSize=12,
        spaceAfter=2
    )

    title_style_3 = ParagraphStyle(
        "address",
        parent=styles["Heading4"],
        fontName="Times-Roman",
        fontSize=12,
        spaceAfter=2
    )

    letter_title_style = ParagraphStyle(
        "Title=Vacation Training",
        fontName="Times-Bold",
        fontSize=12,
        alignment=1,
        leading=13

    )

    salutation_style = ParagraphStyle(
        "Salutation",
        fontName="Times-Roman",
        fontSize=12,
        alignment=0
    )

    date_style = ParagraphStyle(
        "Date",
        fontName="Times-Roman"
    
    )

    body_style = ParagraphStyle(
    "Body",
    fontName="Times-Roman",
    fontSize=12,
    alignment=4,
    leading=15 
    )

    signature_style = ParagraphStyle(
        "Yours_sincerely",
        fontName="Times-Roman",
        fontSize=12,
        alignment=0
    )

    footer_style = ParagraphStyle(
        "Footer",
        fontName="MAIAN",
        fontSize=8,
        alignment=1
    )


    elements = []

    # ======================
    # HEADER (LOGO + TITLE)
    # ======================

    logo_path = os.path.join(BASE_DIR, "source_files", "assets", "logo", "Logo.png")
    logo = Image(logo_path, width=2*cm, height=2*cm)

    header_text = Paragraph(
    """
    <b><font color="red">COLLEGE OF ENGINEERING</font></b>
    <br/>
    <b><font size="10"> KWAME NKRUMAH UNIVERSITY OF SCIENCE AND TECHNOLOGY</font></b>
    <br/>
    <font color="red" size="11">Office of the Provost</font>
    """,
    title_style
    )

    contact_info = Paragraph(
    """
    University Post Office<br/>
    Kumasi-Ghana<br/><br/>
    West Africa<br/>
    Direct Line: 233-3220-60240<br/>
    Tel/Fax: 233-3220-60317<br/>
    Email: provost.coe@knust.edu.gh
    """,
    title_style_2
    )

    header_table = Table([
        [logo, header_text ]
    ], colWidths=[2*cm, 12*cm,], rowHeights=[2*cm], hAlign="LEFT")



    elements.append(header_table)
    elements.append(Spacer(1,12))



    header_table_2 = Table([
        [contact_info]
    ], colWidths=[7*cm], rowHeights=[3*cm], hAlign="RIGHT")

    elements.append(header_table_2)
    elements.append(Spacer(1,0))




    # ======================
    # REFERENCE + DATE
    # ======================

    our_reference = Paragraph("Our Ref: CoE-VACT/2026",title_style_2_1)

    our_date = Paragraph(" Date: <b>Friday, 17th April 2026.</b>", date_style)
    rd = Table([[our_reference,"", our_date]], colWidths=[7*cm,4.8*cm, 7*cm],hAlign="CENTER")
    rd.setStyle(TableStyle([
        ("ALIGN", (0,0), (0,0), "LEFT")
    ]))

    elements.append(rd)
    elements.append(Spacer(1,8))



    # ======================
    # LETTER BODY
    # ======================

    elements.append(Paragraph("Dear Sir/Madam,", salutation_style))
    elements.append(Spacer(1,10))

    elements.append(Paragraph(
    "<u>"
    "<b>VACATION TRAINING FOR 2025/2026 ACADEMIC YEAR</b><br/> "
    "<b>LETTER OF INTRODUCTION</b>"
    "</u>",
    letter_title_style
    ))

    body = f"""
    The College of Engineering at Kwame Nkrumah University of Science and Technology aspires to establish
    itself as a leading global engineering institution committed to national industrial development.

    A vital part of this vision involves providing students with early exposure to industrial practices,
    which is crucial to their training. Therefore, as part of the graduation requirements for our Bachelor
    of Science degree programmes, students must complete a minimum of a 4-week industrial attachment with
    either a local or international industrial organisation between 25th April 2026 and 25th May 2026.
    The aim of this internship, among other objectives, is to enable students to translate theoretical knowledge
    gain in the classroom into practical skills in a real-world work environment.<br/>

    This correspondence aims to express our gratitude to the support you have provided over the years in training
    emerging engineers, as well as your effort in monitoring and evaluating our students during their
    vacation training period.<br/>

    The College of Engineering at KNUST wishes to request a vacation training placement for
    {name}, a third-year student with student ID: {student_id},
    pursuing a Bachelor of Science in {programme} engineering.
    The student can be directly contacted on +233 (0){phone}.<br/>

    We would greatly appreciate your consideration in offering this student a vacation training
    position within your esteemed organisation.<br/>

    Thank you in advance for your cooperation.
    """

    elements.append(Paragraph(body, body_style))
    elements.append(Spacer(1,10))

    # ======================
    # SIGNATURES
    # ======================

    elements.append(Paragraph("Yours sincerely,", signature_style))
    elements.append(Spacer(1,5))

    sig1_path = os.path.join(BASE_DIR, "source_files", "assets", "signatures", "signature1.png")
    sig2_path = os.path.join(BASE_DIR, "source_files", "assets", "signatures", "signature2.png")

    sig1 = Image(sig1_path, width=4*cm, height=2*cm)
    sig2 = Image(sig2_path, width=4*cm, height=2*cm)

    sign_table = Table([
    [
    sig1,
    "",
    sig2
    ]], colWidths=[7*cm,4*cm,7*cm])

    elements.append(sign_table)
    elements.append(Spacer(1,10))


    end1 = Paragraph(
        "Ing. Dr Bright Yeboah-Akowuah<br/>"
        "(College Internship Coordinator)<br/>"
        "<b>+233 (0) 240728535</b>"
    , signature_style)

    end2 =  Paragraph(
        "Ing. Prof Kwabena Biritwum Nyarko<br/>"
        "(Provost, College of Engineering)",signature_style
    )


    end1_2 = Table([[end1,"",end2]], colWidths=[7*cm, 4*cm, 7*cm])

    end1_2.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP")
    ]))

    elements.append(end1_2)
    elements.append(Spacer(1,40))

    # ======================
    # FOOTER
    # ======================

    footer = Paragraph(
    """
    <font color="green">
    Programmes: Aerospace Engineering, Agricultural and Biosystems Engineering,
    Automobile Engineering, Biomedical Engineering, Chemical Engineering,
    Civil Engineering, Computer Engineering, Electrical and Electronic Engineering,Geological Engineering,
    Geomatic Engineering, Industrial Engineering, Marine Engineering, Materials Engineering,
    Mechanical Engineering, Metallurgical Engineering, Petrochemical Engineering, Petroleum Engineering,
    Telecommunication Engineering.
    </font><br/>
    <br/>
    <font color="green">
    Centres: Technology Consultancy Centre, The Energy Centre.
    </font>
    """,
    footer_style
    )

    elements.append(footer)

    # ======================
    # BUILD PDF
    # ======================

    pdf.build(elements, onFirstPage=draw_red_line)
    

    filename = f"{name}_vacation_training.pdf"

    print(f"Saved: {filename}")

    return pdf_path


EMAIL_SENDER = "vacationtrainingknust.coe.office@gmail.com"
EMAIL_PASSWORD = "qhdx ollr awng pvti"

SUBJECT = "VACATION TRAINING LETTER"
BODY_TEMPLATE = """\
Dear Master {name},

Please find attached your internship letter.

Yours faithfully,
Lamech Nartey
"""


    

def attach_files(msg, filepath):
    with open(filepath, 'rb') as file:
        file_data = file.read()
        file_name = os.path.basename(filepath)
    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)


def send_email(receiver_email, receiver_name, attachment_path=None):
    msg = EmailMessage()
    msg["From"] = EMAIL_SENDER
    msg["To"] = receiver_email
    msg["Subject"] = SUBJECT

    msg.set_content(BODY_TEMPLATE.format(name=receiver_name))

    if attachment_path and os.path.exists(attachment_path):
        attach_files(msg, attachment_path)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com",465, context=context) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg) 
        print(f"Email Sent to {receiver_name}")


app = Flask(__name__)
app.secret_key = "secret123"



@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form.get("name")
        email = request.form.get("email")
        student_id = request.form.get("student_id")
        programme = request.form.get("programme")
        phone = request.form.get("phone")

        if not all([name, email, student_id, programme, phone]):
            flash("All fields are required")
            return redirect(url_for("index"))

        # Generate PDF
        pdf_file = generate_letter(name, student_id, programme, phone)

        # Send email
        try:
            send_email(email, name, pdf_file)
        except Exception as e:
            print(e)
            flash("Letter generated but email could not be sent.")
            return redirect(url_for("index"))

        return send_file(pdf_file, as_attachment=True)

    return render_template("form.html")



