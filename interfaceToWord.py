from flask import Flask
from flask import request, send_file
interfaceToWord = Flask(__name__)
application = interfaceToWord # our hosting requires application in passenger_wsgi

from docx import Document
#from StringIO import StringIO
from io import BytesIO
from bs4 import BeautifulSoup

import justext
from convert import processElement

@interfaceToWord.route("/", methods=['GET','POST'])
def hello():
    html = request.values.get('htmlContent',default="", type=str)

#    paragraphs = justext.justext( html, justext.get_stoplist("English"))

    if html=="":
        return "No HTML content detected\n"

    doc = Document()
    soup = BeautifulSoup( html, 'html.parser')

    for elem in soup.children:
        elemType=processElement(elem,doc)

    f = BytesIO()
    doc.save(f)
    f.seek(0)
    return send_file(f,attachment_filename='convert.docx',
            as_attachment=True , mimetype='text/docx')

if __name__ == "__main__":
    interfaceToWord.run(debug=True,ssl_context='adhoc')


