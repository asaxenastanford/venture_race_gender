import xlrd
import openpyxl
import requests
import os, shutil

def run(command):
    print(command)
    exit_status = os.system(command)
    if exit_status > 0:
        exit(1)

javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx'
wbJ = xlrd.open_workbook(javaFile)
javasheetread = wbJ.sheet_by_index(0)

javaFile1 = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/names_to_run_new.xlsx'
wbJ1 = xlrd.open_workbook(javaFile1)
namesread = wbJ1.sheet_by_index(0)

xfile1 = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')
javasheetwrite = xfile1['Sheet1'] 


for r in range(500, 1668):
	index = round(namesread.cell_value(r,2))
	#print(round(index))
	male_count = namesread.cell_value(r,3)
	female_count = namesread.cell_value(r,4)
	final_gender = namesread.cell_value(r,5)
	eth_count = namesread.cell_value(r,6)
	eth = namesread.cell_value(r,7)

	male_index = "G" + str(index)
	female_index = "H" + str(index)
	#print(male_index)
	#print(str(male_count))
	javasheetwrite[male_index] = male_count
	javasheetwrite[female_index] = female_count
	
	if(male_count > 0 or female_count > 0):
		if(eth_count != ""):
			confidence = float(eth_count)/float(male_count + female_count)
			confidence_index = "S" + str(index)
			javasheetwrite[confidence_index] = confidence
		else:
			javasheetwrite[genderIndex] = "NA"
		genderIndex = "I" + str(index) 
		javasheetwrite[genderIndex] = final_gender
	else:
		javasheetwrite[genderIndex] = "NA"
		
	eth_count_index = "J" + str(index)
	eth_index = "K" + str(index)		
	race_index = "T" + str(index)
	sourceIndex = "U" + str(index) 
	javasheetwrite[eth_count_index] = eth_count
	if(len(eth) > 0):
		javasheetwrite[eth_index] = eth
		javasheetwrite[race_index] = eth
	else:
		javasheetwrite[eth_index] = "NA"
		javasheetwrite[race_index] = "NA"	
		
	javasheetwrite[sourceIndex] = "face_plus_plus"

	if(r%25==0):
		print(str(r))
		xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')

xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/name_prism_processing/Venture_Capital_List_Mod.xlsx')