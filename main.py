import os    # library for operation system
import comtypes.client   # library for MS Word 
import docx               # library for work with docx files

def convert_docx_to_pdf(docx_path, pdf_output_path):   
    word = comtypes.client.CreateObject("Word.Application")  # create object in MS Word
    word.Visible = False                               # app MS Word is not visiable 

    in_file = word.Documents.Open(docx_path)    # open docx file

    in_file.ExportAsFixedFormat(pdf_output_path, ExportFormat=17)  # export file to pdf. 17 is for format PDF

    in_file.Close()   # close the file and the app MS Word 
    word.Quit()
    print("PDF soubory vytvořeny.")

def main():   # HERE WE START... 
    try:
        word_path = input("Zadejte cestu ke složce s docx soubory: ")    # user input path to folder with docx files 
        pdf_path = input("Zadejte adresu složky, kam se uloží PDF soubory: ")  # path for the folder with the new files 

        for root, dirs, files in os.walk(word_path):  # cycle for every file in folder and subfolder 
            for doc_file in files:
                if doc_file.endswith(".docx"):   # if file is docx, cycle continue 
                    docx_path = os.path.join(root, doc_file)   # create the path for the file 
                    pdf_output_path = os.path.join(pdf_path, f"{os.path.splitext(doc_file)[0]}.pdf")  # create the path for the new pdf file 
           
                    convert_docx_to_pdf(docx_path, pdf_output_path)   # finally, the convertion 
        else:
            print("Chyba. Ve složce nejsou žádné docx soubory. Zkontrolujte její cestu.")
    except Exception as e:
        print(f"Chyba {e}.")
   
main()  # ACTUALLY HERE WE START... 




