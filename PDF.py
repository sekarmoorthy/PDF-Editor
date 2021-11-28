file1 = ""
import tkinter as tk
import io,os,time
from tkinter import messagebox
from tkinter import *
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import docx
from fpdf import FPDF
text1 = ""
root = tk.Tk()
text = tk.Text(root,width=100,height=40)
root.title("Editing Window")
def load_file(pages=None):
    print("\n                      [!] close the Window(Pop-up) after editing The document to save it")
    file1 = input("\n                                            Enter the File Name : ")
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)
    try:  
        file_taken =open(file1,'rb')
        with file_taken as data:
            output = io.StringIO()
            manager = PDFResourceManager()
            converter = TextConverter(manager, output, laparams=LAParams())
            interpreter = PDFPageInterpreter(manager, converter)
            text.grid()
            for page in PDFPage.get_pages(file_taken, pagenums):
                interpreter.process_page(page)
            file_taken.close()
            converter.close()
            text1 = output.getvalue()
            output.close
            text.insert(1.0,text1)   
    except FileNotFoundError:
        print("File Name Incorrect or File Is not present There.... :(")       
def con_to_word():   
    file_name = input("\n                                            Enter the File Name : ")
    if(file_name == ""):
        print("\n                                            Please Enter File Name .....!!!!!")
        time.sleep(1.5)
        file_name = input("\n                                            Enter the File Name : ")
    original_file_path = file_name.replace(os.sep, '/')
    base_name = os.path.basename(file_name)
    time.sleep(1.5)
    print("\n                                            Starting to Converted...!!") 
    time.sleep(1.5)
    pages=None
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)
    mydoc = docx.Document()
    output = io.StringIO()
    manager = PDFResourceManager()
    print("\n                                            Converting....!!!")
    time.sleep(1.5)
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)
    try:
        for page in PDFPage.get_pages(open(original_file_path,'rb'), pagenums):
            interpreter.process_page(page)
        text1 = output.getvalue()
        changed_xml = ''.join(c for c in text1 if valid_xml_char_ordinal(c))
        mydoc.add_paragraph(changed_xml)
        pathing = len(base_name)
        mydoc.save(file_name[:-pathing]+"Converted "+base_name[:-4]+".docx")
        time.sleep(1.5)
        print("\n                                            Successfully Converted")
        time.sleep(1.5)
        print("\n                                            File Saved as : 'Converted as "+base_name[:-4]+".txt'")
        time.sleep(1.5)
    except FileNotFoundError:
        print("File Name Incorrect or File Is not present There.... :(")
    os.system('cls')
def con_to_text():
    file_name = input("\n                                            Enter the File Name : ")
    if(file_name == ""):
        print("\n                                            Please Enter File Name .....!!!!!")
        time.sleep(1.5)
        file_name = input("\n                                            Enter the File Name : ")
    base_name = os.path.basename(file_name)
    original_file_path = file_name.replace(os.sep, '/')
    extract_file_path = file_name.replace('/', os.sep)
    time.sleep(1.5)
    print("\n                                            Starting to Convert...!!") 
    time.sleep(1.5)
    pages=None
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)
    output = io.StringIO()
    print("\n                                            Converting....!!!")
    time.sleep(1.5)
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)
    try:
        for page in PDFPage.get_pages(open(original_file_path,'rb'), pagenums):
            interpreter.process_page(page)
        text1 = output.getvalue()
        changed_xml = ''.join(c for c in text1 if valid_xml_char_ordinal(c))
        pathing = len(base_name)
        with open(file_name[:-pathing]+"Converted "+base_name[:-4]+".txt","w",encoding='utf-8') as f:
            f.write(changed_xml)
        print("\n                                            Successfully Converted")
        time.sleep(1.5)
        print("\n                                            File Saved as : 'Converted as "+base_name[:-4]+".txt'")
        time.sleep(1.5)
        wish = input("\n                                            Do you want to open the Text Document You Converted ....![y/n]\n\n                                            ")
        if(wish == "y"):
            print("\n                                            Opening Document...!")
            os.system("notepad "+extract_file_path[:-pathing]+"Converted "+base_name[:-4]+".txt")
            print("\n                                            Reverting Changes...!")
            print("\n                                            Happy Converting...!")
            time.sleep(2)
        else:
            print("\n                                            Reverting Changes...!")
            print("\n                                            Happy Converting... :)")
            time.sleep(2)
    except FileNotFoundError:
        print("\n                                            File Name Incorrect or File Is not present There.... :(")
        time.sleep(4.5)
    os.system('cls')
def valid_xml_char_ordinal(c):
    codepoint = ord(c)
    return( 
           0x20 <= codepoint <= 0xD7FF or codepoint in (0x9, 0xA, 0xD) or 0xE000 <= codepoint <= 0xFFFD or 0x10000 <= codepoint <= 0x10FFFF)      
def on_closing():
    base_name = os.path.basename(file1)
    extract_file_path = file1.replace('/', os.sep)
    pathing = len(base_name)
    if messagebox.askokcancel("Alert", "Do you want to Save as Pdf & quit?"):
        input_values = text.get("1.0",'end-1c')
        changed_xml = ''.join(c for c in input_values if valid_xml_char_ordinal(c))
        with open("prev_file.txt","w",encoding='utf-8') as f:
            f.write(changed_xml)
        pdf = FPDF()   
        pdf.add_page()
        pdf.set_font("Arial", size = 12)
        with open('prev_file.txt', 'r') as file :
            filedata = file.read()
        filedata = filedata.replace('•', '->')
        with open('file.txt', 'w') as file:
            file.write(filedata)
        f = open("prev_file.txt", "r",encoding='latin-1')
        for x in f:
            pdf.cell(150, 10, txt = x, ln = 1, align = 'J')
        pdf.output(name="converted file.pdf",dest=extract_file_path[:-pathing])
        root.destroy()
        print("\n                                            File Saved as : 'Converted file.pdf'")
        time.sleep(3)
        f.close()
        file.close()
        os.remove("prev_file.txt")
        os.remove("file.txt")
        os.system('cls')
switch = {1:'edit pdf and save it',2:'save pdf to Word',3:'save pdf to text'}
def loop():
    os.system('cls')
    print(r'''
                                                         ________
                                                        |________|
                                        ____     __  _____ |  |   ____ _______
                                      _/ __ \ | |  \   |   |  | _/ __ \\_  __ \
                                       \  __/ | |  |   |   |  |  \ ___/ |  | \/
                                        \___  | |__/   |   |__|         |__|
                                                     -----
      ''')
    switch_value = int(input("                                                                         [!] Don't Close the Window that Poped-up\n\n                                            1:edit pdf and save it\n                                            2:save pdf to Word\n                                            3:save pdf to text\n                                            4:Exit the Program\n                                            \n                                            Enter any option : "))
    if(switch_value == 1):
        root.after_idle(load_file)
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    elif(switch_value == 2):
        con_to_word()
    elif(switch_value == 3):
        con_to_text()
    elif(switch_value == 4):
        os.system('exit')
    elif(switch_value == 5):
        os.system('cls')
        os.system("pip install PyPDF2 fpdf docx")
        time.sleep(2)
for i in range(1,100):
    loop()
