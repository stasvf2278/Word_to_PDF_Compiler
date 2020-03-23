import os
from win32com import client
import time
import PyPDF2
wdFormatPDF = 17

def main():

    folder = r'C:\Users\stanm\Desktop\Word_to_PDF_Compiler'                         ## Change root directory (file path) here

    folderPath = os.path.normpath(folder)

    file_type = 'docx'
    out_folder = folderPath + "\\PDF"

    try:                                                                            ## Converts .docx to .pdf
        word = client.DispatchEx("Word.Application")
        for files in os.listdir(folder):
            if files.endswith(".docx") or files.endswith('doc'):
                out_name = files.replace(file_type, r"pdf")
                in_file = os.path.abspath(folderPath + "\\" + files)
                out_file = os.path.abspath(folderPath + "\\" + out_name)
                doc = word.Documents.Open(in_file)
                print ('Exporting', out_file)
                doc.SaveAs(out_file, FileFormat=17)
                doc.Close()
    except Exception as e: print ('e')

    finally:
        word.Quit()                                                                 ## Closes .docx file

    # Get all the PDF filenames.

    pdfFiles = []                                                                   #Becomes list of new PDFs
    for filename in os.listdir('.'):                                                #Iterates through to find PDFs
        if filename.endswith('.pdf'):
            pdfFiles.append(filename)
    pdfFiles.sort(key=str.lower)                                                    #Sorts PDFs by order

    print(pdfFiles)                                                                 #Lists PDFs in order

    pdfWriter = PyPDF2.PdfFileWriter()

    for filename in pdfFiles:                                                       #Appends final PDF document by iteration
        pdfFileObj = open(filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        for pageNum in range(0, pdfReader.numPages):                                #Loop to append page by page 0 in range(0 - starts on page 1
            pageObj = pdfReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

    pdfOutput = open('OUTPUT.pdf', 'wb')                                   ## Compiles and saves thesis to single document - Change output name here
    pdfWriter.write(pdfOutput)
    pdfOutput.close()

main()