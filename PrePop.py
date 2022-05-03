#Import tkinter library for GUI interface
from math import trunc
import tkinter
from tkinter.filedialog import askdirectory, askopenfilename
#Import tkinter for excel manipulation
import xlrd
#Import pdfrw for pdf manipulation
import pdfrw
from datetime import datetime

#May or may not define a constructor 
def __init__(self, temp_dir, output_dir):
    self.temp_directory = temp_dir
    self.output_dir = temp_dir

try:
    #GUI for XL worksheet
    xl_source = askopenfilename(title="Choose the completed Excel Form")
    #Opening worksheet with XLRD
    xl_file = xlrd.open_workbook(xl_source)

    #GUI for output directory 
    output_dir = askdirectory(title="Choose your desired output directory")
    #print(output_dir)

    #Take the sheet object from the workbook to access # of columns and rows 
    sheet = xl_file.sheets()
    #Take the first sheet of the XL document (only one present)
    sheet = sheet[0]
    #Initialize dictionary
    col_dict = {}
    #Iterate through the columns 
    for row in range(sheet.ncols):
        #Initialize the list that will contain the contents of each columns 
        temp_l = []
        #Iterate through the rows
        for col in range(sheet.nrows):
            #The first row is the header -> need to be Key not Value
            if col != 0:
                temp_l.append(sheet.cell_value(col,row))
        #Create the key-value pair
        if col != 0:
            col_dict[(sheet.cell_value(0,row))] = temp_l
        #Initialize the keys 
        else:
            col_dict[(sheet.cell_value(0,row))] = None
    #print(col_dict)

    #Scrape out the fillable fields from the PDF 
    pdf_temp = pdfrw.PdfReader(r'C:\Users\svc_grp_training\OneDrive - The Equitable Life Insurance Company of Canada\Documents\Python Scripts\test.pdf')
    #Iterate through all of the pages of the PDF document
    for page in pdf_temp.pages:
        #Take out all editable fields
        blanks = page['/Annots']
        #Check that annotations are instantiated 
        if blanks is None:
            continue

        #Iterate through the names of said fields 
        for blank in blanks:
            if blank['/Subtype'] == '/Widget':
                try:
                    key = blank['/T'][1:-1]
                    #print("KEY",key)

                    for headings in col_dict:
                        i = 0 
                        #print(headings)

                        if headings.lower() in key.lower() and headings != "Address":
                            print(headings,key)
                            print(col_dict[headings][i])
                            if type(col_dict[headings][i]) == float:
                                print( col_dict[headings][i])
                                col_dict[headings][i] = trunc(col_dict[headings][i])
                                print( col_dict[headings][i])
                            pdfstr = pdfrw.objects.pdfstring.PdfString.encode(str(col_dict[headings][i]))
                            blank.update(pdfrw.PdfDict(V=pdfstr))
                            col_dict[headings].pop(i)
                            break

                        elif key == "date2":
                            pdfstr = pdfrw.objects.pdfstring.PdfString.encode(datetime.today().strftime('%m/%d/%Y'))
                            blank.update(pdfrw.PdfDict(V=pdfstr))
                            col_dict[headings].pop(i)
                            break
                            
                            #print(headings, key)
                                
                    
                    # for heading in col_dict:
                    #     i = 0 
                    #     #print("This is the heading",heading)
                    #     #print("This is the key", key)
                    #     if heading.lower() in key.lower():
                    #         print(key, heading)
                    #         #print("This is the key", key, "This is the dictionary elem" ,col_dict[heading][i])
                    #         #PLACE THE ACTUAL ELEMENT 
                    #         col_dict[heading].pop(i)
                    #         #print(col_dict[heading][i])
                    #         break  
                #If there is a NoneType we want to catch the error so we can skip over it
                except TypeError:
                    continue

        pdf_temp.Root.AcroForm.update(
            pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        pdfrw.PdfWriter().write(output_dir + "data.pdf", pdf_temp)


#If a file is not selected or an .xls file is not selected then an error box is displayed.
except FileNotFoundError:
    tkinter.messagebox.showerror(title="Error", message="A file was not selected. The program will close.")
except xlrd.biffh.XLRDError:
    tkinter.messagebox.showerror("You did not select the appropriate Excel Document.")