from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def fill_pdf(names,codes,admin):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    nameYstart = 728
    codeYstart = 710
    for i in range(len(names)):
        yPos = nameYstart-(i*136)
        can.drawString(140,yPos,names[i])
        can.drawString(455,yPos,admin)
        can.drawString(90,codeYstart-(i*136),codes[i])

    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open("roster.pdf", "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open("destination.pdf", "wb")
    output.write(outputStream)
    outputStream.close()