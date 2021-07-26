#writeCheck.py
#check for distresses, write to XML if present

#colored text
colorama.init(autoreset=True)

#load input/workbook selected from main.py and set current active sheet to "sheet"
workbook = load_workbook(filename=db_name, data_only=True)
workbook.sheetnames
sheet = workbook.active

#for printing:
#set/show last row and column to avoid out of range error
#subtracts data starts at 0, LastRow+1 to show actual rows read starting at 1 and not 0
LastRow=sheet.max_row-RowIncr
print("Total row number is ", LastRow+1)

#for actual cell checking: data starts at 0, LastRow for readability/printing above, FinalRow for data crawling from data start. 
#bizarre idiosyncracies in openpyxl: I required two end row variables to find a blank row, compare and set the last row with legible data below
FinalRow = sheet.max_row

#check to see if last row has data for first column - if 0 then subtract 1 from last Row
#to avoid empty data/date error
#else, add LastRow to RowIncr(sheet actual-data-start offset)

#DETERMINE IF FURTHER ROWS ARE BLANK/0 AND stop pulling data - 
#had issues with openpyxl concerning dates and blank rows, which required this odd fix to find the blank row at end of spreadsheet
cellCheck=sheet.cell(row=FinalRow, column=1).value
if cellCheck == 0:
	LastRow=LastRow+RowIncr-1
else:
	LastRow=LastRow+RowIncr


#cut filename of .xml/.log file to 4 char for date
fileName=db_name[0:4]
filesRun.append(fileName)

#create folders for xml and log files
if not os.path.exists('XML'):
    os.makedirs('XML')

if not os.path.exists('LOG'):
    os.makedirs('LOG')

#write xml and log files if not currently present
if os.path.exists("XML/"+fileName+".xml"):
  os.remove("XML/"+fileName+".xml")
if os.path.exists("LOG/"+fileName+".log"):
  os.remove("LOG/"+fileName+".log")
  print(Style.DIM + Fore.RED +"Original XML/LOG Deleted \n \n" + Style.RESET_ALL + Fore.RESET)
else:
  print(Style.DIM + Fore.RED +"New file created \n \n")
f = open("XML/"+fileName + ".xml", "a+")
logFile = open("LOG/"+fileName + ".log", "a+")


#add headers to top of xml file
print(xml_header, "\n", xml_schema, sep="", file=f)

#how many rows read/written
rowsRead=0
#ticker used to determine if a distress was located in the row - set to 1 for distress located, defaulted to 0
ticker = 0


#information previously required for logging the info where data failed to read/write
def emptyData():
	print ("Empty: ", row[INSPECTED_PID2],":", row[SAMPLENUMBER], " ###################", file=logFile)
	#print ("Empty: ", row[INSPECTED_PID2],":", row[SAMPLENUMBER], " ###################")

def fullData():
		print("Data: ", row[INSPECTED_PID2], ":", row[SAMPLENUMBER], file=logFile)
		#print("Data: ", row[INSPECTED_PID2], ":", row[SAMPLENUMBER])

		
#check each code section to ensure there is a distress for this row, set ticker to 1 to write that distress to XML
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

#loading bar which iterates through the codeCheck on each distress to check for distresses and then write to XML
#I attempted to iterate through a list created with all the distress_codes, but too many errors were present, it was simpler to just run the codeCheck function
#as opposed to continue to debug and iterate through list/dictionary with the code names.
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
			




	
#close xml main tag and workbook
print("</pavementData>", sep="", file=f)
f.close()
logFile.close()

#print rows and files read
print(Fore.MAGENTA + "\nRows Read: ", int(rowsRead))		
print(Fore.MAGENTA + "File Opened:", db_name)
print(Fore.MAGENTA + "Files Written:", fileName+".xml"+", "+fileName+".log" + "\n\n")
