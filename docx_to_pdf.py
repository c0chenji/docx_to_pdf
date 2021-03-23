import sys
from PDFNetPython3 import *
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import white, black
import docx


class Docx:
    def __init__(self, filename):
        self.filename = filename

    def read(self):
        try:
            doc = docx.Document(self.filename)

        except:
            print("Could not open/read file:{}".format(self.filename))
            sys.exit()
        return doc.paragraphs


def create_pdf(fname):
    c = canvas.Canvas(fname)
    c.setFont("Courier", 12)
    return c


def convert_to_pdf(paragraphs, input_list, output_file):
    
    c = create_pdf(output_file)
    form = c.acroForm

    # Initiate x, y coordinates
    start_x = 20
    start_y = 650

    label_lst = []
    label = ""

    # 
    form_length = 0

    input_index = 0
    for paragraph in paragraphs:
        content = paragraph.text
        if len(content) > 0:
            for i in range(len(content)):
                # start calculating required form size when "_" occurs
                if content[i] == "_":
                    form_length += 1
                    # draw form para in pdf
                    if i+1 < len(content):
                        # draw form and reset form_length
                        if content[i+1] == " ":
                            try:
                                form.textfield(value=str(input_list[input_index]), name=label_lst[-1], tooltip=label_lst[-1],
                                               x=start_x, y=start_y, borderStyle='inset',
                                               borderColor=black, fillColor=white,
                                               width=form_length*12, height=17,
                                               textColor=black, forceBorder=True)
                                start_x += form_length*12+10
                                # label_lst.append("blank:{}".format(form_length))
                                form_length = 0
                                input_index += 1
                                # print(label_lst[-1])
                            except:
                                print("Something goes wrong with input ")
                    # Reach the end of paragraph, create input field
                    else:
                        try:
                            form.textfield(value=str(input_list[input_index]), name=label_lst[-1], tooltip=label_lst[-1],
                                           x=start_x, y=start_y, borderStyle='inset',
                                           borderColor=black, fillColor=white,
                                           width=form_length*12, height=17,
                                           textColor=black, forceBorder=True)
                            # update next x-pos
                            input_index += 1
                            start_x += form_length*12+10
                            # print(label_lst[-1])
                        except:
                            print("Something goes wrong with input ")
                # skip if it is a space
                elif content[i] == " ":
                    continue

                else:
                    label = label + content[i]
                    if i+1 < len(content):
                        if content[i+1] == " " or content[i+1] == "_":
                            c.drawString(start_x, start_y, label)
                            start_x += len(label)*8
                            label_lst.append(label)
                            label = ""
            # reset start_x
            start_x = 20
        else:
            label_lst.append("  ")
        # start a new line
        start_y -= 30
    c.drawString(start_x,start_y,"Signature: ")    
    c.save()
    #start a new line for signature field
    start_y -= 20
    add_signature_field(output_file,start_x,start_y)

def add_signature_field(outpath,x,y):
    doc = PDFDoc(outpath)
    page1 = doc.GetPage(1)
    # Create a new signature form field in the PDFDoc. The name argument is optional;
    # leaving it empty causes it to be auto-generated. However, you may need the name for later.
    # Acrobat doesn't show digsigfield in side panel if it's without a widget. Using a
    # Rect with 0 width and 0 height, or setting the NoPrint/Invisible flags makes it invisible.
    certification_sig_field = doc.CreateDigitalSignatureField("Signiture")
    widgetAnnot = SignatureWidget.Create(
        doc, Rect(x, y, x+150, y-50), certification_sig_field)
    page1.AnnotPushBack(widgetAnnot)

    # (OPTIONAL) Add an appearance to the signature field.
    # img = Image.Create(doc.GetSDFDoc(), appearance_image_path)
    # widgetAnnot.CreateSignatureAppearance(img)

    # Prepare the document locking permission level to be applied upon document certification.
    certification_sig_field.SetDocumentPermissions(
        DigitalSignatureField.e_annotating_formfilling_signing_allowed)

    # Save the PDFDoc. Once the method below is called, PDFNet will also sign the document using the information provided.
    doc.Save(outpath, 0)

if __name__ == "__main__":
    # test1 single line form
    #input_list [name, age]
    target1 = Docx("test1.docx").read()
    convert_to_pdf(target1, ["John", 30], "output1.pdf")

    # test2
    #input_list [name, age, phone, school]
    target2 = Docx("test2.docx").read()
    convert_to_pdf(target2, ["John", 30, 123123, "blabla"], "output2.pdf")
