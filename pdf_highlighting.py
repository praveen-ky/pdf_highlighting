import fitz
import sys
import pandas as pd

def excel():
    Excel_name = input("Enter the name of the Excel Sheet :")
    Sheet_name = input("Enter the Sheet Name :")
    df = pd.read_excel(Excel_name,sheet_name=Sheet_name)

    # Storing in a list format by taking a column
    tcs = df['EMPLOYEE NAME'].tolist()
    print(tcs)

    print("Total no. of Elements in the Columns:",len(tcs))

    # Opening the Pdf File and storing the Data in doc
    pdf_file = input("Enter the name of the PDF File: ")
    doc = fitz.open(pdf_file)

    for count in range(len(doc)):
        multi_instances=[]
        page = doc[count]

        for person in range(len(tcs)):         
            quads = page.searchFor(str(tcs[person]), hit_max=100, quads=True)

            for instance in quads:
                highlight = page.addHighlightAnnot(instance)

    #Saving the File
    doc.save("Output1final.pdf",garbage=4, deflate=True, clean=True)
