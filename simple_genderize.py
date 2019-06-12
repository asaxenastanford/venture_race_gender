from genderize import Genderize
import xlrd
import openpyxl
import requests
import os, shutil

# to search 
excelFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/Venture_Capital_List_Mod.xlsx'
wb = xlrd.open_workbook(excelFile) 
sheet = wb.sheet_by_index(0) #read
xfile = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/Venture_Capital_List_Mod.xlsx')
sheet2 = xfile['Sheet1'] #xfile.get_sheet_by_name('Sheet1') #write

xfile1 = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/names_to_run.xlsx')
javasheetwrite = xfile1['Sheet1'] #xfile.get_sheet_by_name('Sheet1')
javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/names_to_run.xlsx'
wbJ = xlrd.open_workbook(javaFile)
javasheetread = wbJ.sheet_by_index(0)

currentLine = javasheetread.cell_value(0, 1)
currentLineNumber = round(currentLine)
total = javasheetread.cell_value(0, 3)
totalNumber = round(total)
api_string = "df689b8f770275905ab667f3ed4b3191"

for r in range(2,32046):
	firstName = sheet.cell_value(r,1)
	lastName = sheet.cell_value(r,2)
	result = (Genderize(api_key = api_string).get([firstName]))
	#print(type(result))
	dictionary = result[0]
	name = dictionary['name']
	gender = dictionary['gender']
	genderIndex = "D" + str(r + 1)
	sheet2[genderIndex] = gender

	if(gender == None):
		firstNameIndex = "A" + str(currentLineNumber)
		lastNameIndex = "B" + str(currentLineNumber)
		actualIndex = "C" + str(currentLineNumber)

		javasheetwrite[firstNameIndex] = firstName
		javasheetwrite[lastNameIndex] = lastName
		javasheetwrite[actualIndex] = (r + 1)
		currentLineNumber += 1
		totalNumber += 1
	else:
		probability = dictionary['probability']
		probIndex = "E" + str(r + 1)
		sheet2[probIndex] = probability

		count = dictionary['count']
		countIndex = "F" + str(r + 1)
		sheet2[countIndex] = count

	if(r%25==0):
		print(dictionary)
		javasheetwrite["B1"] = currentLineNumber
		javasheetwrite["D1"] = totalNumber
		xfile.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/Venture_Capital_List_Mod.xlsx')
		xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/names_to_run.xlsx')

javasheetwrite["B1"] = currentLineNumber
javasheetwrite["D1"] = totalNumber

xfile.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/Venture_Capital_List_Mod.xlsx')
xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/genderize/names_to_run.xlsx')
