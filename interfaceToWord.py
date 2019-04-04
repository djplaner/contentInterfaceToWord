from flask import Flask
from flask import request
interfaceToWord = Flask(__name__)
application = interfaceToWord # our hosting requires application in passenger_wsgi

@interfaceToWord.route("/")
def hello():
    html = request.args.get('htmlContent',default="", type=str)

    if html=="":
        return "no content was passed\n"
    else:
        return "Content was passed \n %s \n" % html

if __name__ == "__main__":
    interfaceToWord.run()


