import sys
import docx
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import white, black


class docx_to_pdf:
    def __init__(self, filename):
        self.filename = filename

    def read(self):
        try:
            doc = docx.Document(self.filename)

        except docx.opc.exceptions.PackageNotFoundError:
            print("Could not open/read file:{}".format(self.filename))
            sys.exit()
        return doc.paragraphs


def create_pdf(fname):
    c = canvas.Canvas(fname)
    c.setFont("Courier", 12)
    return c


target = docx_to_pdf("test.docx").read()

def add_form(input, na):
    c = create_pdf("test_output.pdf")
    form = c.acroForm
    start_x = 20
    start_y = 650

    label_lst = []
    label = ""

    form_length = 0

    index_value = 0
    for paragraph in input:
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
                            form.textfield(value=str(na[index_value]), name=label_lst[-1], tooltip=label_lst[-1],
                                           x=start_x, y=start_y, borderStyle='inset',
                                           borderColor=black, fillColor=white,
                                           width=form_length*12, height=17,
                                           textColor=black, forceBorder=True)
                            start_x += form_length*12+10
                            # label_lst.append("blank:{}".format(form_length))
                            form_length = 0
                            index_value += 1
                            print(label_lst[-1])
                    # Reach the end of paragraph, create input field
                    else:
                        form.textfield(value=str(na[index_value]), name=label_lst[-1], tooltip=label_lst[-1],
                                       x=start_x, y=start_y, borderStyle='inset',
                                       borderColor=black, fillColor=white,
                                       width=form_length*12, height=17,
                                       textColor=black, forceBorder=True)
                        # update next x-pos
                        index_value += 1
                        start_x += form_length*12+10
                        print(label_lst[-1])

                # skip if it is a space
                elif content[i] == " ":
                    continue

                else:
                    label = label + content[i]
                    if i+1 < len(content):
                        if content[i+1] == " " or content[i+1] == "_":
                            c.drawString(start_x, start_y, label)
                            start_x += len(label)*12
                            label_lst.append(label)
                            label = ""
            # reset start_x
            start_x = 20
        else:
            label_lst.append("  ")
        # start new ling
        start_y -= 30
    c.save()


add_form(target, ["join", 32,"out"])
# doc = docx.Document("test.docx")
# print(type(doc.paragraphs[0].text))
