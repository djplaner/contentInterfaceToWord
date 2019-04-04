from flask import Flask
from flask import request, send_file
interfaceToWord = Flask(__name__)
application = interfaceToWord # our hosting requires application in passenger_wsgi

from docx import Document
#from StringIO import StringIO
from io import BytesIO

import justext

@interfaceToWord.route("/", methods=['GET','POST'])
def hello():
    html = request.values.get('htmlContent',default="", type=str)

    paragraphs = justext.justext( html, justext.get_stoplist("English"))

    doc = Document()

#    ty = type(paragraphs)
#    doc.add_paragraph( ty )

#    doc.add_paragraph( paragraphs[2] )

    output = ""
    count=0
    for p in paragraphs: 
        doc.add_paragraph( p.text)
#        print("paragraph is %s" % paragraph )
#        output+="************************* %s" %count
#        count+=1
#        output+=paragraph.text

#    p = doc.add_paragraph( "Hello world from Word land")


    if html=="":
        return "no content was passed\n"
    else:
        f = BytesIO()
        #f.write('Hello world')
        doc.save(f)
        f.seek(0)
        return send_file(f,attachment_filename='convert.docx',
                as_attachment=True , mimetype='text/docx')

if __name__ == "__main__":
    interfaceToWord.run(debug=True)


