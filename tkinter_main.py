# aplikace s uživatelským rozhraním vytvořeným v knihovně Tkinter. 
# zapotřebí jsou dva obrázky: docx_picture_png.png a pdf_picture_png.png

import os    # library for operation system
import comtypes.client   # library for MS Word 
from docx import Document
import time
import base64
from tkinter import *
from tkinter import filedialog
from tkinter import PhotoImage
import sys
from pathlib import Path
from PIL import Image, ImageTk

def convert_docx_to_pdf(docx_path, pdf_output_path):   
    try:
        word = None
        in_file = None
        word = comtypes.client.CreateObject("Word.Application")  # create object in MS Word
        word.Visible = False                               # app MS Word is not visiable 

        in_file = word.Documents.Open(docx_path)    # open docx file

        in_file.ExportAsFixedFormat(pdf_output_path, ExportFormat=17)  # export file to pdf. 17 is for format PDF

        in_file.Close()   # close the file and the app MS Word 
        word.Quit()
        return True
        
    except Exception as e:
        pass 
        print(f"Stala se chyba při otevírání dokumentu {e}")
        return False
    
 
def main():   # HERE WE START... 
    update_status_label("Makám na tom, vydržte chvilku prosím...") 
    successful_conversions = 0
    word_path = word_folder_entry.get()  #  input("Zadejte cestu ke složce s DOCX soubory: ")    # user input path to folder with docx files

    while not os.path.isdir(word_path) or not any(file.endswith(".docx") for file in os.listdir(word_path)): # check user input
        update_status_label("Zadali jste neplatnou cestu nebo ve složce nejsou žádné DOCX soubory.")          
        return         

    pdf_path = pdf_folder_entry.get() # ("Zadejte adresu složky, kam se uloží PDF soubory: ")  # path for the folder with the new files 
    while not os.path.isdir(pdf_path):       # check user input
        update_status_label("Zadali jste neplatnou cestu.")
        return   
 
    for root, dirs, files in os.walk(word_path):  # cycle for every file in folder and subfolder 
        for doc_file in files:
            if doc_file.endswith(".docx"):   # if file is docx, cycle continue
                docx_path = os.path.abspath(os.path.join(root, doc_file))
 
#                docx_path = os.path.join(root, doc_file)   # create the path for the file 
                pdf_output_path = os.path.join(pdf_path, f"{os.path.splitext(doc_file)[0]}.pdf")  # create the path for the new pdf file 
        
                if convert_docx_to_pdf(docx_path, pdf_output_path):   # finally, the convertion 
                    successful_conversions += 1
    
    update_status_label(f"Úspěšně bylo převedeno {successful_conversions} soubory/ů DOCX do PDF.")
    thank_label.config(text="Děkujeme za použití aplikace.")
    
  
def browse_word_folder():        # browse by user
    folder_path = filedialog.askdirectory()
    word_folder_entry.delete(0, 'end')
    word_folder_entry.insert(0, folder_path)

def browse_pdf_folder():
    folder_path = filedialog.askdirectory()
    pdf_folder_entry.delete(0, "end")
    pdf_folder_entry.insert(0, folder_path)

def update_status_label(new_text):       # status label update
    status_label.config(text=new_text)
    window.update_idletasks()         # actualization gui


# barvy
main_color = "#14085f"
foot_label_color = "#5C8C46"
button_color = "#BDE038"

# okno
window = Tk()
window.minsize(800, 400)
window.resizable(False, False)              # not able to make bigger or smaller by user´s mouse
window.title("Aplikace na převod DOCX do PDF")
window.config(bg=main_color)                 # background color

# obrázky 
docx_img = PhotoImage(file="docx_picture_png.png")
pdf_img = PhotoImage(file="pdf_picture_png.png")

# řádka, kam se vypíše adresa docx 
word_folder_entry = Entry(window, width=0)

# řádka, kam se vypíše adresa pdf 
pdf_folder_entry = Entry(window, width=1)

# tlačítko procházet složky uživatelem - docx
word_folder_button = Button(window, text="Vyberte složku", font=("Arial", 12), command=browse_word_folder, image=docx_img, compound="bottom", width=170, height=240)
word_folder_button.grid(row=0, column=0, padx=20, pady=20)

# tlačítko procházet složky uživatelem - pdf 
pdf_folder_button = Button(window, text="Vyberte složku", font=("Arial", 12), command=browse_pdf_folder, image=pdf_img, compound="bottom", width=170, height=240)
pdf_folder_button.grid(row=0, column=2, padx=20, pady=20)

# převod
convert_button = Button(window, text="Spustit převod!", font=("Arial", 12), command=main, bg=button_color, width=15, height=5)
convert_button.grid(row=0, column=1)

# label, který vypisuje hlášky o činnosti aplikace, co zrovna dělá a zda je úspěšná
status_label = Label(window, text="", font=("Arial", 12))
status_label.grid(row=1, column=1, columnspan=1, padx=10, pady=(10))

# patička
foot_label = Label(window, text="APP created by Pavel Kopecký & Jan Pevný & Jan Stehlík.", font=("Arial", 14), bg=foot_label_color)
foot_label.grid(row=2, column=1, columnspan=1)

# label hláška a poděkování
thank_label = Label(window, text="", font=("Arial", 12))
thank_label.grid(row=3, column=1, columnspan=1, pady=5, padx=5)

# hlavní cyklus
window.mainloop()



