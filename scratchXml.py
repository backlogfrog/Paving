
#custom XML module YATTAG wouldn't take data correctly so I said Fuck it and XML will be written from scratch
#Distress comment for every distress listed
DistressComment = "Imported"
#row[INSPECTED_SIZE]-- set actual square footage - excel sheet has different size
#inspectedSize=row[P_LENGTH]*row[P_WIDTH]
inspectedSize = row[SAMPLESIZE]
#PCI condition, doesn't import
condition=1

#Spacing for xml
iS="  "

#set pid without editing excel sheet to format NETWORK::STREET::PID#
#addressID can only be 10 characters
address=str(row[INSPECTED_PID1]).replace(" ","")
address=address.replace(".","")
address=address.replace("/","")
address=address.upper()
#cut address to 10 char
addressCut=address[0:10]
fullpid="WINFIELD::" + addressCut + "::" + str(row[INSPECTED_PID2])



#set date to proper format

spread_date = row[INSPECTED_DATE]
parsed_date = datetime.strftime(spread_date, "%m/%d/%Y")

#temporary date needed due to date change and not present in excel
#or set dateSet=parsed_date to pull a proper parsed date from spreadsheet
dateSet=parsed_date


#manually write XML data to filename.xml
#Changed comment to say Imp: for imported to identify

print(iS, "<geospatialInspectionData inspectionDate=\"",dateSet,"\"", " units=\"English\" level=\"SAMPLE\" >", sep="", file=f)
print(iS*2, "<inspectedElement inspectedElementID=\"", row[SAMPLENUMBER], "\"", " size=\"", inspectedSize, "\" ", "PID=\"", fullpid, "\" inspectedType=\"R\" comment=\"",DistressComment,  "\" noDistresses=\"false\">",  sep="", file=f)


#if row[ALLIGATOR_S] > 0 or row[POTHOLE_S] > 0:
print (iS*3, "<inspectionData>", sep="", file=f)



#Set distress codes to dict, check to see if ANY codes are > 0, then write
#iterating through isn't working...
#unneeded at this time, but I don't have to type row[blahblah] now...
#initially attempted to run through dictionary checking if > 0 for each distress, iterating through was throwing errors for a list or dict on if > 0 - works at this time!

distressCodes = {
	'sweatherC': row[SWEATHERING_CODE],
	'sweatherS': row[SWEATHERING_S],
	'sweatherQ': row[SWEATHERING_Q],
	'alligatorC': row[ALLIGATOR_CODE],
	'alligatorS': row[ALLIGATOR_S],
	'alligatorQ': row[ALLIGATOR_Q],
	'blockcrackC': row[BLOCKCRACK_CODE],
	'blockcrackS': row[BLOCKCRACK_S],
	'blockcrackQ': row[BLOCKCRACK_Q],
	'trasnverseC': row[TRANSVERSE__CODE],
	'trasnverseS': row[TRANSVERSE_S],
	'trasnverseQ': row[TRANSVERSE_Q],
	'depressionC': row[DEPRESSION_CODE],
	'depressionS': row[DEPRESSION_S],
	'depressionQ': row[DEPRESSION_Q],
	'potholeC': row[POTHOLE_CODE],
	'potholeS': row[POTHOLE_S],
	'potholeQ': row[POTHOLE_Q],
	'edgecrackC': row[EDGECRACKING_CODE],
	'edgecrackS': row[EDGECRACKING_S],
	'edgecrackQ': row[EDGECRACKING_Q],
	'jointspallC': row[JOINTSPALLING_CODE],
	'jointspallS': row[JOINTSPALLING_S],
	'jointspallQ': row[JOINTSPALLING_Q],
	'durabilityC': row[DURABILITYCRACKING_CODE],
	'durabilityS': row[DURABILITYCRACKING_S],
	'durabilityQ': row[DURABILITYCRACKING_Q],
	'faultC': row[FAULTING_CODE],
	'faultS': row[FAULTING_S],
	'faultQ': row[FAULTING_Q],
	'patchingC': row[PATCHING_CODE],
	'patchingS': row[PATCHING_S],
	'patchingQ': row[PATCHING_Q],
	'bumpsagC': row[BUMPSAG_CODE],
	'bumpsagS': row[BUMPSAG_S],
	'bumpsagQ': row[BUMPSAG_Q]
}

#used for checking data as we run through
distressCheck = []
distressCheck += distressCodes.values()
#print (distressCheck)





#Function for printing distress to file
def distressPrint(code, severity, quantity):
	try:
		if float(code) > 0:
			print (iS*6, "<levelDistress distressCode=\"", code, "\"", " severity=\"", severity, "\" quantity=\"", abs(quantity), "\"", " comment=\"", DistressComment, "\" />", sep="", file=f)
			fullData()
	except ValueError:
		emptyData()


#forced to run function individually with each var unfortunately: iterating through the dict/list was not working, int/str issues with "> 0" check to ensure we're not printing empty distresses. 
print (iS*5, "<PCIDistresses>", sep="", file=f)
distressPrint(distressCodes["sweatherC"],distressCodes["sweatherS"],distressCodes["sweatherQ"])
distressPrint(distressCodes["alligatorC"],distressCodes["alligatorS"],distressCodes["alligatorQ"])
distressPrint(distressCodes["blockcrackC"],distressCodes["blockcrackS"],distressCodes["blockcrackQ"])
distressPrint(distressCodes["trasnverseC"],distressCodes["trasnverseS"],distressCodes["trasnverseQ"])
distressPrint(distressCodes["depressionC"],distressCodes["depressionS"],distressCodes["depressionQ"])
distressPrint(distressCodes["potholeC"],distressCodes["potholeS"],distressCodes["potholeQ"])
distressPrint(distressCodes["edgecrackC"],distressCodes["edgecrackS"],distressCodes["edgecrackQ"])
distressPrint(distressCodes["jointspallC"],distressCodes["jointspallS"],distressCodes["jointspallQ"])
distressPrint(distressCodes["durabilityC"],distressCodes["durabilityS"],distressCodes["durabilityQ"])
distressPrint(distressCodes["faultC"],distressCodes["faultS"],distressCodes["faultQ"])
distressPrint(distressCodes["patchingC"],distressCodes["patchingS"],distressCodes["patchingQ"])
distressPrint(distressCodes["bumpsagC"],distressCodes["bumpsagS"],distressCodes["bumpsagQ"])
print (iS*5, "</PCIDistresses>", sep="", file=f)


#close initial opened xml tags
#if row[ALLIGATOR_S] > 0 or row[POTHOLE_S] > 0:
print (iS*3, "</inspectionData>", sep="", file=f)
print(iS*2, "</inspectedElement>", sep="", file=f)
print(iS, "</geospatialInspectionData>", sep="", file=f)

#generic "inspection info," seems to be needed for each inspection despite appearance of duplicate information. 
print (iS, "<geospatialInspectionData level=\"SECTION\" units=\"English\" inspectionDate=\"", dateSet, "\" >", sep="", file=f)
print (iS*2, "<inspectedConditions PID=\"", fullpid, "\" >", sep="", file=f)
print (iS*3, "<conditions>", sep="", file=f)
print (iS*4, "<levelCondition comment=\"\" source=\"\" cndMeasureUID=\"StructPCI\" cndMeasure=\"SCI\" conditionText=\"\" condition=\"", condition, "\"/>", sep="", file=f)
print (iS*3, "</conditions>", sep="", file=f)
print (iS*2, "</inspectedConditions>", sep="", file=f)
print (iS, "</geospatialInspectionData>", sep="", file=f)


#end