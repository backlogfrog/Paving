#writeCheck.py
#check for distresses, write to XML if present

colorama.init(autoreset=True)

#load input/workbook from above and set current active sheet to "sheet"
workbook = load_workbook(filename=db_name, data_only=True)
workbook.sheetnames
sheet = workbook.active


#set/show last row and column to avoid out of range error
#LastRow=sheet.max_row
#LastColumn=sheet.max_column
#print ("Last Column ", LastColumn)


#something is fucky here, last row shows is greater than 
LastRow=sheet.max_row-RowIncr
print("Total row number is ", LastRow+1)




#LastRow = Copy of 2012.xlsx 685 readable lines
#2012.xlsx has 686 readable lines
FinalRow = sheet.max_row

#check to see if last row has data for first column - if 0 then subtract 1 from last Row
#to avoid empty data/date error
#else, add LastRow to RowIncr(sheet actual-data-start offset)

#DETERMINE IF FURTHER ROWS ARE BLANK/0 AND stop pulling data, im not sure how this works currently 2.19.20
cellCheck=sheet.cell(row=FinalRow, column=1).value
if cellCheck == 0:
	LastRow=LastRow+RowIncr-1
else:
	LastRow=LastRow+RowIncr


#cut filename of .xml/.log file to 4 char
fileName=db_name[0:4]
filesRun.append(fileName)

#create folders for xml and log files
if not os.path.exists('XML'):
    os.makedirs('XML')

if not os.path.exists('LOG'):
    os.makedirs('LOG')

#write files/delete if present
if os.path.exists("XML/"+fileName+".xml"):
  os.remove("XML/"+fileName+".xml")
if os.path.exists("LOG/"+fileName+".log"):
  os.remove("LOG/"+fileName+".log")
  print(Style.DIM + Fore.RED +"Original XML/LOG Deleted \n \n" + Style.RESET_ALL + Fore.RESET)
else:
  print(Style.DIM + Fore.RED +"New file created \n \n")
f = open("XML/"+fileName + ".xml", "a+")
logFile = open("LOG/"+fileName + ".log", "a+")






#add headers to top of page
print(xml_header, "\n", xml_schema, sep="", file=f)

#how many rows read/written
rowsRead=0
ticker = 0

def emptyData():
	print ("Empty: ", row[INSPECTED_PID2],":", row[SAMPLENUMBER], " ###################", file=logFile)
	#print ("Empty: ", row[INSPECTED_PID2],":", row[SAMPLENUMBER], " ###################")

def fullData():
		print("Data: ", row[INSPECTED_PID2], ":", row[SAMPLENUMBER], file=logFile)
		#print("Data: ", row[INSPECTED_PID2], ":", row[SAMPLENUMBER])

#check each code to ensure there is a distress for this row, set ticker to 1 to write
def codeCheck(code):
	
	try:
		if float(code) > 0:
			global ticker
			ticker = 1
			fullData()
			
		else:
			emptyData()
			
	except ValueError:
		
		
		emptyData()
		ticker = 0


with alive_bar(LastRow-1, bar = 'bubbles', spinner = 'dots_waves') as bar:
	for row in sheet.iter_rows(min_row=RowIncr, max_row=LastRow, values_only=True):
			rowsRead=rowsRead+1
			ticker = 0
			codeCheck(row[SWEATHERING_CODE])	 	
			codeCheck(row[ALLIGATOR_CODE])
			codeCheck(row[BLOCKCRACK_CODE])
			codeCheck(row[TRANSVERSE__CODE])
			codeCheck(row[DEPRESSION_CODE])
			codeCheck(row[POTHOLE_CODE])
			codeCheck(row[EDGECRACKING_CODE])
			codeCheck(row[JOINTSPALLING_CODE])
			codeCheck(row[DURABILITYCRACKING_CODE])
			codeCheck(row[FAULTING_CODE])
			codeCheck(row[PATCHING_CODE])
			codeCheck(row[BUMPSAG_CODE])
			if ticker == 1:
				exec(open("scratchXml.py").read())
			bar()
			#bar.next()
			



#print (Fore.RED + "FIRST SS IS IN 2010 - NO SET DATE \n \n")
	
#close xml main tag and workbook
print("</pavementData>", sep="", file=f)
f.close()
logFile.close()

#print rows and files read
print(Fore.MAGENTA + "\nRows Read: ", int(rowsRead))		
print(Fore.MAGENTA + "File Opened:", db_name)
print(Fore.MAGENTA + "Files Written:", fileName+".xml"+", "+fileName+".log" + "\n\n")