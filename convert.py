
from docx import Document
from io import BytesIO
from bs4 import BeautifulSoup
import bs4

#import justext
import re

from doc import html

#---------------------------------------------------------------
# handleTable( table, doc )
# - given a bs4 table element, add it into the document

def handleTable( table, doc ):

    table_body = table.find('tbody')

    docTable = doc.add_table()

    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')

#---------------------------------------------------------------
# Given an element from BeautifulSoup identify which type of
# element is and figure out how to add it to the Word document

def processElement( elem, doc ):

    #-- ignore new lin
    if isinstance( elem, bs4.element.NavigableString):
        if elem.string == '\n':
            pass
        else:
            print( repr(elem) )
            print("SOMETHING STRANGE")
    else:
        tag = elem.name

        if tag == 'div':
            pass
        elif ( tag == 'ol'  ):
            pass
        elif ( tag == 'li'  ):
            pass
        elif ( tag == 'h1' ):
            doc.add_heading( elem.contents )
        elif ( tag == 'h2' ):
            doc.add_heading( elem.contents, 2 )
        elif ( tag == 'h3' ):
            doc.add_heading( elem.contents, 3 )
        elif ( tag == 'p' ):
            print( "P\n %s" % elem.contents )
            numChildren = len(elem.contents )
            if ( numChildren==1 ):
                doc.add_paragraph( elem.contents)
            else:
                #-- have to figure out how to handle nested elements
                #   -- paragraphs and runs
                pass

        elif ( tag == 'span' ):
            pass
        elif ( tag == 'table' ):
            ## docx tables https://python-docx.readthedocs.io/en/latest/api/table.html
            ## bs4 tables 
            # https://roche.io/2016/05/scrape-wikipedia-with-python
            handleTable( elem, doc )
            pass
        elif ( tag == 'em' ):
            ## for em, span and other inline tags, will need to
            #  convert them and then add them back to the paragraph and continue
            #  the next para ***HARD**
            pass
        elif ( tag == 'blockquote' ):
            pass
        else:
            print( "ERROR can't handle %s" % tag )
            print( repr(elem))

soup = BeautifulSoup( html, 'html.parser')

#print( soup.prettify() )

doc = Document()

for elem in soup.children: #soup.next_siblings:
#    print("---------------------------------")
#    print( repr(elem) )
#    print(type(elem))

    elemType = processElement(elem,doc)
#    print("xxxxxxxxxxxxxxxxxxxx-------------")


doc.save("word.docx")

#---- justext version
#paragraphs = justext.justext( html, justext.get_stoplist("English"))

#for p in paragraphs:
#    print(p.text)

#    doc = Document()

#    doc.add_paragraph( html ) 
#    ty = type(paragraphs)
#    doc.add_paragraph( ty )

#    doc.add_paragraph( paragraphs[2] )

#    for p in paragraphs: 
#        doc.add_paragraph( p)
#        print("paragraph is %s" % paragraph )
#        output+="************************* %s" %count
#        count+=1
#        output+=paragraph.text

#    p = doc.add_paragraph( "Hello world from Word land")


