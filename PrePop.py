#Truncation for number excel values
from math import trunc
#Import tkinter library for GUI interface
import tkinter
from tkinter.filedialog import askopenfilename
#Import tkinter for excel manipulation
import xlrd
#Import pdfrw for pdf manipulation
import pdfrw
#Time function and file management 
from datetime import datetime
import os

pol_dict = ["name of policyholder","name of plan sponsor"]
cert_dict = ["certificateno","certificate number","institution code","cert#"]
dob_dict = ["date of birth mmddyyyy","birthdate", ""]
date_dict = ["date signed", "date2"]
name_dict = ["plan member's name", "plan members name first middle initial last","full name of primary beneficiary first middle last" ]
emp_dict = ["date employed full time mmddyyyy","date employed full time" ]


def PdfCreator():
    try:
        root = tkinter.Tk()
        root.wm_attributes('-topmost', 1)
        root.withdraw()

        #GUI for XL worksheet
        xl_source = askopenfilename(parent=root, title="Choose the completed Excel Form", filetypes=[("Excel Workbooks",".xls")])
        #Opening worksheet with XLRD
        xl_file = xlrd.open_workbook(xl_source)

        #GUI for reference PDF location
        output_dir = askopenfilename(parent=root,title="Choose PDF reference form",filetypes=[("PDF",".pdf")])

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

        #Scrape out the fillable fields from the PDF 
        pdf_temp = pdfrw.PdfReader(output_dir)
        cum_string = ""
        counter = 0 

        #Iterate through all of the pages of the PDF document
        while len(col_dict.get("Policyholder")) != 0:
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
                            key = key.lower()

                            for headings in col_dict:
                                i = 0 
                                #Looking for similar names and matching 
                                if headings.lower() in key  and not (headings == "Address" and key == "email address"):
                                    cum_string = col_dict[headings][i]
                                    #Number format in XL shows up as a flot. Remove all decimal values
                                    if type(col_dict[headings][i]) == float:
                                        cum_string = str(trunc(col_dict[headings][i]))
                                    #Break and move to the next element
                                    break
                                
                                elif headings == "Policyholder" and key  in pol_dict:
                                    cum_string = col_dict[headings][i]
                                    break

                                elif headings == "Certificate #" and key  in cert_dict:
                                    cum_string = col_dict[headings][i]
                                    break
                                
                                elif headings == "Date of Employment" and key in emp_dict:
                                    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + trunc(col_dict[headings][i]) - 2).strftime('%m/%d/%Y')
                                    dt = str(dt).split()
                                    cum_string = str(dt[0])
                                    break

                                #Concatenating the names into the singular field
                                elif (key  in name_dict and ("name" in headings.lower() and " member" in headings.lower())):
                                    cum_string += col_dict[headings][i] + " "
                        
                                    if headings == "Plan Member Last Name":
                                        cum_string = cum_string.strip()
                                        #Break and move to the next element
                                        break

                                    else:
                                        del col_dict[headings][i]

                                #Number of hours
                                elif key  == "Number of hours worked per week" and headings == "Standard Hours":
                                    cum_string = str(col_dict[headings][i])
                                    break

                                #HCSA Value
                                elif key == "undefined" and headings == "HCSA":
                                    cum_string = str(col_dict[headings][i])
                                    break
                                
                                #Edge case considering the final date window. Initialize with todays date and time
                                elif key  in date_dict:
                                    cum_string = datetime.today().strftime('%m/%d/%Y')
                                    break

                                # elif len(col_dict[headings]) != 0: 
                                #     if key == "Check Box02" and col_dict[headings][i] == "English":
                                #         val_str = pdfrw.objects.pdfname.BasePdfName('/Yes')
                                #         blank.update(pdfrw.PdfDict(V=val_str))
                                #         break

                                #     elif key == "Check Box200" and col_dict[headings][i] == "French":  
                                #         val_str = pdfrw.objects.pdfname.BasePdfName('/Yes')
                                #         blank.update(pdfrw.PdfDict(V=val_str))
                                #         break

                            #Only want to fill in spaces if we have a value in the excel sheet 
                            if cum_string != "":
                                #Update the field inside of the PDF 
                                pdfstr = pdfrw.objects.pdfstring.PdfString.encode(cum_string)
                                blank.update(pdfrw.PdfDict(V=pdfstr))
                                #Remove the element from the list so it does not get repeated 
                                #Since the current date is not a column we do not want to delete an element
                                if key  not in date_dict:
                                    del (col_dict[headings][i])
                                #Reset the cumulative string
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

    os.startfile(cleaned_out)

PdfCreator()