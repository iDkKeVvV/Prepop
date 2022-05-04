#Import tkinter library for GUI interface
from math import trunc
import tkinter
from tkinter.filedialog import askopenfilename
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

    #GUI for reference PDF location
    output_dir = askopenfilename(title="Choose PDF reference form")

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
    pdf_temp = pdfrw.PdfReader(output_dir)

    cum_string = ""
    #while(col_dict.get())
    #Iterate through all of the pages of the PDF document
    while len(col_dict.get("Policyholder")) != 0:
        counter = 0 
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

                        for headings in col_dict:
                            i = 0 
                            #Looking for similar names and matching 
                            if headings.lower() in key.lower() and not (headings == "Address" and key == "email address"):
                                cum_string = col_dict[headings][i]
                                #Number format in XL shows up as a flot. Remove all decimal values
                                if type(col_dict[headings][i]) == float:
                                    cum_string = str(trunc(col_dict[headings][i]))
                                #Break and move to the next element
                                break

                            elif headings == "Certificate #" and key == "cert#":
                                cum_string = col_dict[headings][i]
                            
                            elif headings == "Date of Employment" and key == "Date Employed Full time mmddyyyy":
                                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + trunc(col_dict[headings][i]) - 2).strftime('%m/%d/%Y')
                                dt = str(dt).split()
                                cum_string = str(dt[0])

                            elif (key =="Plan Members Name first middle initial last" and ("name" in headings.lower() and " member" in headings.lower())):
                                cum_string += col_dict[headings][i] + " "
                                
                                if headings == "Plan Member Last Name":
                                    cum_string = cum_string.strip()
                                    #Break and move to the next element
                                    break

                            elif key == "Number of hours worked per week" and headings == "Standard Hours":
                                cum_string = str(col_dict[headings][i])
                                break
                            
                            elif key == "undefined" and headings == "HCSA":
                                cum_string = str(col_dict[headings][i])
                                break
                            
                            #Edge case considering the final date window. Initialize with todays date and time
                            elif key == "date2":
                                cum_string = datetime.today().strftime('%m/%d/%Y')
                                break
                        
                        if cum_string != "":
                            print("What we are deleteing" ,headings)
                            print(len(col_dict["Policyholder"]))
                            #Update the field inside of the PDF 
                            pdfstr = pdfrw.objects.pdfstring.PdfString.encode(cum_string)
                            blank.update(pdfrw.PdfDict(V=pdfstr))
                            #Remove the element from the list so it does not get repeated 
                            #print(key)
                            del (col_dict[headings][i])
                            cum_string = ""

                    #If there is a NoneType we want to catch the error so we can skip over it
                    except TypeError:
                        continue
            
        #Update the PDF so that the filled elements are visible from the start.
        pdf_temp.Root.AcroForm.update(
            pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        #Write the new PDF to the location specified earlier
        #print(output_dir)
        split_out = output_dir.split('/')
        split_out.pop(-1)
        cleaned_out ='/'.join(split_out)
        #print("This is cleaned out ", cleaned_out)
        pdfrw.PdfWriter().write(cleaned_out + "/filled_" + str(counter) + ".pdf", pdf_temp)
        counter += 1


#If a file is not selected or an .xls file is not selected then an error box is displayed.
except FileNotFoundError:
    tkinter.messagebox.showerror(title="Error", message="A file was not selected. The program will close.")
except xlrd.biffh.XLRDError:
    tkinter.messagebox.showerror(title="Error", message="You did not select the appropriate Excel Document.")
except pdfrw.errors.PdfParseError:
    tkinter.messagebox.showerror(title="Error", message="You did not select the appropriate PDF Document.")