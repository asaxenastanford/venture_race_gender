import xlrd
import openpyxl
import requests
import os, shutil

def run(command):
    print(command)
    exit_status = os.system(command)
    if exit_status > 0:
        exit(1)

javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/gender_eval.xlsx'
wbJ = xlrd.open_workbook(javaFile)
javasheetread = wbJ.sheet_by_index(0)

ventureBook = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx'
v = xlrd.open_workbook(ventureBook)
ventureread = v.sheet_by_index(0)

xfile1 = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')
javasheetwrite = xfile1['Sheet1'] 

nextPersonRow = 2
nextPersonCol = 0 
nextPerson = round(javasheetread.cell_value(nextPersonRow,nextPersonCol))  
print("Next: " + str(nextPerson))

for r in range(2, 32046):
	gender = ventureread.cell_value(r,19)
	source = ventureread.cell_value(r,20)
	currentPerson = round(ventureread.cell_value(r,0))
	bert = ""

	if(currentPerson == nextPerson):
		gender = javasheetread.cell_value(nextPersonRow,23)
		source = javasheetread.cell_value(nextPersonRow,24)
		bert = javasheetread.cell_value(nextPersonRow,21)
		nextPersonRow += 1
		nextPerson = round(javasheetread.cell_value(nextPersonRow,nextPersonCol))

	confidenceIndex = "Z" + str(r + 1)
	genderIndex = "AA" + str(r + 1)
	sourceIndex = "AB" + str(r + 1)

	javasheetwrite[genderIndex] = gender
	javasheetwrite[sourceIndex] = source

	confidence = 0

	if(source == "genderize"):
		confidence = ventureread.cell_value(r,4)
	elif(source == "faceplusplus"):
		male_count = ventureread.cell_value(r,6)
		female_count = ventureread.cell_value(r,7)

		if(male_count > 0 or female_count > 0):
			confidence = float(male_count)/float(male_count + female_count)
			if(confidence < 0.5):
				confidence = 1 - confidence
	elif(source == "nameprism"):
		confidence = ventureread.cell_value(r,18)
	else:
		confidence = 1


	bertIndex = "V" + str(r + 1)
	javasheetwrite[bertIndex] = bert
	javasheetwrite[confidenceIndex] = confidence

	if(r%25==0):
		print(str(r))
		xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')

xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')