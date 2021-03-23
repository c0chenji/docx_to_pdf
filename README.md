## Description
A feature in Python that reads a form as the word document and generates the form in PDF format.  This feature will be able to insert field value into the PDF.  This feature will also be able to digitally sign the PDF. 
## Requirements
Python version: 3.7.10.

**Note**:
[PDFNetPython3](https://www.pdftron.com/blog/python/python3/) only supports Python 3.5 - 3.8 on Windows (x64, x86), Linux (x64, x86), and macOS.  
Use [Conda](https://docs.conda.io/projects/conda/en/latest/user-guide/tasks/manage-environments.html#removing-an-environment) to create a virtual environment and install the required python version.

## Packages installed 

```bash
certifi==2020.12.5
lxml==4.6.3
PDFNetPython3==8.1.0
Pillow==8.1.2
python-docx==0.8.10
reportlab==3.5.66
wincertstore==0.2

```

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install required libraries.

```bash
pip install -r requirements.txt
```
## Test
**test1.docx** and **test2.docx** should be created
```bash
python docx_to_pdf.py
```