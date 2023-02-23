from doctest import master
import pdfrw
import os 
#The path here was removed for security reasons
directory = os.fsencode(path)

wrd_dict = {}
master_list = []

for file in os.listdir(directory):
    
    try:
        wrk_fl = os.fsdecode(file)
        print(wrk_fl)
        if wrk_fl.endswith('.pdf') and "FR" not in wrk_fl:
            string_nm = path + str(wrk_fl)
            pdf_temp = pdfrw.PdfReader(string_nm)
            cum_string = ""
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

                            if "check" in key.lower() or len(key) > 50:
                                continue
                            
                            if key.lower() not in wrd_dict:
                                wrd_dict[key.lower()] = 1
                            
                            else:
                                wrd_dict[key.lower()] += 1
                                
                        #If there is a NoneType we want to catch the error so we can skip over it
                        except TypeError:
                            continue

    except pdfrw.errors.PdfParseError:
        continue
    except TypeError:
        continue

for keys in wrd_dict:
    if(wrd_dict[keys] > 0):
        master_list.append((wrd_dict[keys],keys))

master_list.sort(reverse=True)

print(master_list)
