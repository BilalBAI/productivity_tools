from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from PyPDF2 import PdfFileWriter, PdfFileReader

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


def generate_water_pdf(content):
    '''
  Generate watermark.pdf
  :param content: watermark name
  :return:
  '''
    cans = canvas.Canvas('watermark.pdf', pagesize=(21 * cm, 29.7 * cm))
    cans.translate(10 * cm, 12 * cm)
    cans.setFont('Helvetica', 30)
    cans.setFillColorRGB(0.5, 0.5, 0.5)
    cans.rotate(45)
    cans.drawString(-7 * cm, 0 * cm, content)
    cans.drawString(7 * cm, 0 * cm, content)
    cans.drawString(0 * cm, 7 * cm, content)
    cans.drawString(0 * cm, -7 * cm, content)
    cans.save()


def insert_water_to_pdf(input_pdf, output_pdf, water_pdf):
    '''
  merge watermakr.pdf to the main pdf
  :param input_pdf: input fild dir
  :param output_pdf: output fild dir
  :param water_pdf: output fild dir
  :return:
  '''
    water = PdfFileReader(water_pdf)
    water_page = water.getPage(0)
    pdf = PdfFileReader(input_pdf, strict=False)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(water_page)
        pdf_writer.addPage(pdf_page)
    output_file = open(output_pdf, 'wb')
    pdf_writer.write(output_file)
    output_file.close()


def send_mail(send_from,
              send_to,
              subject,
              text,
              files=None,
              server="127.0.0.1"):
    assert isinstance(send_to, list)

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(fil.read(), Name=basename(f))
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()


# client email: pdf doc dir (without .pdf)
config = {
    'test1@test.com': 'C:/test/test1',
    'test2@test.com': 'C:/test/test2',
    'test3@test.com': 'C:/test/test3'
}
temp_watermark_dir = 'C:/test/watermark.pdf'

if __name__ == '__main__':
    for i in config.keys():
        generate_water_pdf(i)
        insert_water_to_pdf(f"{config[i]}.pdf", f"{config[i]}_output.pdf",
                            temp_watermark_dir)
        send_mail('sophia@test.com', i, 'TEST', 'test',
                  f"{config[i]}_output.pdf")
