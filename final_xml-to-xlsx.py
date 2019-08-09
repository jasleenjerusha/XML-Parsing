import xml.etree.cElementTree as ET
import openpyxl
from openpyxl import Workbook
from datetime import datetime

def read_chapters(book, blist):
    for chapters in book.findall('chapter_id'):
        chapid     =  chapters.find('id').text
        chaptname  =  chapters.find('name').text
        dbname     =  chapters.find('dbname').text
        score      =  chapters.find('score').text 
        blist.append([" ", chapid, chaptname, dbname, score])
        print ("{:8}{:25}{:20}{:15}{:10}".format(" ", chapid, chaptname, dbname, score))
 
def read_book(book, blist):
    bid       =  book.find('book_id').text
    name      =  book.find('name_of_the_book').text
    btype     =  book.find('book_type').text
    chapid    =  book.find('chapter_id').text
    print ("{:8}{:25}{:20}{:15}".format(bid, name, btype, chapid.strip()))
    blist.append([bid, name, btype, chapid.strip()])
    read_chapters(book, blist)
    print ()

    return

def read_xml(xml_file, blist):
    tree = ET.ElementTree(file=xml_file)
    root = tree.getroot()
    for book in tree.iter(tag='book'):
        read_book(book, blist)
   
def init_xlsx(bname):
    titles = ["bookid", "Name/chapter", "Type/Chapt", "DB-Name", "score", "Location"]
    wb = Workbook(bname)
    wsheet = wb.create_sheet(index=1, title="raghu")
    wsheet.append(titles)
 
    return (wb, wsheet)


def write_to_xlsx(wsheet, book_data):
    print ()

    for data in book_data:
        wsheet.append(data)

def main():
    book_name = "book_data.xlsx"
    book_data = []
    wb, wsheet = init_xlsx(book_name)
    print ("{:8}{:25}{:15}{:15}{:10}{:15}".format("bookid", "Name/chapter", "Type/Chapt", "DB-Name", "score", "Location"))

    read_xml("data.xml", book_data)
    write_to_xlsx(wsheet, book_data)
    wb.save(book_name)
    return

if (__name__ =="__main__"):
    main()
