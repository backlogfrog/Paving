#extract data from a spreadsheet and convert to Paver XML data for import

#finding files by extension in main dir/opening files, exec .py files
import os
from os import system, name
import glob

def clear(): 
	# for windows 
	if name == 'nt':
		_ = system('cls')
	# for mac and linux(here, os.name is 'posix')
	else: 
		_ = system('clear') 

#justprogressbarthings
from alive_progress import alive_bar
#from progress.bar import Bar


#needed to parse date in scratchXml.py
from datetime import datetime

#colored text for funskies while I wait for data
import colorama
from colorama import Fore, Back, Style 
colorama.init()
#reset color/style after printing
colorama.init(autoreset=True)
#Colors, yay...
#Format: print(Fore.BLACK + "TEXT " + str(VARIABLE))
#Fore: BLACK, RED, GREEN, YELLOW, BLUE, MAGENTA, CYAN, WHITE, RESET.
#Back: BLACK, RED, GREEN, YELLOW, BLUE, MAGENTA, CYAN, WHITE, RESET.
#Style: DIM, NORMAL, BRIGHT, RESET_ALL

#openpyxl package opens excel sheets
from openpyxl import load_workbook

#NEW DATA HAS NO DATE - NEED TO SET DATE

##################################################################################
##################################################################################
##################################################################################
#set start positions to read data: 2012.xlsx data starts at row 4
#RowIncr=4
RowIncr=2   
##################################################################################
##################################################################################
##################################################################################
##################################################################################
##################################################################################




#import class info from inspectionClasses.py for inspection data and from mapping.py based on column to be printed to XML

from mapping import INSPECTED_SIZE, INSPECTED_DATE, INSPECTED_PID1, INSPECTED_PID2, DCOMMENT, P_LENGTH, P_WIDTH, SAMPLENUMBER, SWEATHERING_CODE, SWEATHERING_S, SWEATHERING_Q, ALLIGATOR_CODE, ALLIGATOR_S, ALLIGATOR_Q, BLOCKCRACK_CODE, BLOCKCRACK_S, BLOCKCRACK_Q, TRANSVERSE__CODE, TRANSVERSE_S, TRANSVERSE_Q, DEPRESSION_CODE, DEPRESSION_S, DEPRESSION_Q, POTHOLE_CODE, POTHOLE_S, POTHOLE_Q, EDGECRACKING_CODE, EDGECRACKING_S, EDGECRACKING_Q, JOINTSPALLING_CODE, JOINTSPALLING_S, JOINTSPALLING_Q, DURABILITYCRACKING_CODE, DURABILITYCRACKING_S, DURABILITYCRACKING_Q, FAULTING_CODE, FAULTING_S, FAULTING_Q, PATCHING_CODE, PATCHING_S, PATCHING_Q, BUMPSAG_CODE, BUMPSAG_S, BUMPSAG_Q, SAMPLESIZE

#set header/schema for xml
xml_header = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>"
xml_schema = "<pavementData xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:noNamespaceSchemaLocation=\"PavementInspectionDataV2.xsd\">"


#############################open workbook from xlsx files in main dir for mult choice
############################# check if user wants to continue to run conversions

filesRun = []


while True:
	print(Fore.BLUE + "░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░")
	print(Fore.BLUE + "░█▀▀ ▀▄▀ █▀▀ █▀▀ █░░   ▀█▀ █▀█   █▀█ ▄▀█ █░█ █▀▀ █▀█░")
	print(Fore.BLUE + "░██▄ █░█ █▄▄ ██▄ █▄▄   ░█░ █▄█   █▀▀ █▀█ ▀▄▀ ██▄ █▀▄░")
	print(Fore.BLUE + "░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░\n\n")
	excel_list = [f for f in glob.glob("*.xlsx")]
	print(Fore.BLUE + "Spreadsheets:" + Fore.RESET)
	for i, db_name in enumerate(excel_list, start=1):
		print ('{}. {}'.format(i, db_name))
	while True:
		try:
			colorama.Fore.RESET
			if filesRun:
				 print("\nREAD:", *filesRun, sep = ":")  
			selected = int(input(Fore.CYAN + 'Select Spreadsheet (1-{}): '.format(i)))
			db_name = excel_list[selected-1]
			print(Style.DIM + 'Loading {}'.format(db_name), "\n")
			exec(open("writeCheck.py").read())
			break
		except (ValueError, IndexError):
			print(Fore.BLUE + 'Invalid. Please enter number between 1 and {}!'.format(i))
			print("	")



	cont = input(Fore.CYAN + "Another? Y/N > ")
	clear()
	while cont.lower() not in ("y","n"):
		cont = input(Fore.CYAN + "Another? Y/N > " + Fore.RESET)
		clear()
	if cont == "n":
		break


