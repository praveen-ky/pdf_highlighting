import fitz
import sys
import pandas as pd

df = pd.read_excel('Powai_Mar_Salary.xlsx',sheet_name='Sheet1')

# Storing in a list format by taking a column
tcs = df['EMPLOYEE NAME'].tolist()
print(tcs)

print("Total no. of Elements in the Columns:",len(tcs))

# Opening the Pdf File and storing the Data in doc
doc = fitz.open('input.pdf')

for count in range(len(doc)):
    multi_instances=[]
    page = doc[count]

    for person in range(len(tcs)):         
        quads = page.searchFor(str(tcs[person]), hit_max=100, quads=True)
        
        for instance in quads:
            highlight = page.addHighlightAnnot(instance)

#Saving the File
doc.save("Output1final.pdf",garbage=4, deflate=True, clean=True)