import xlrd
import openpyxl
import requests
import os, shutil

def run(command):
    print(command)
    exit_status = os.system(command)
    if exit_status > 0:
        exit(1)

javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/cs224u/gender_eval.xlsx'
wbJ = xlrd.open_workbook(javaFile)
javasheetread = wbJ.sheet_by_index(0)

xfile1 = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/cs224u/gender_eval.xlsx')
javasheetwrite = xfile1['Sheet1'] 
differences = 0

for r in range(2, 1101):
	male_count = javasheetread.cell_value(r,6)
	female_count = javasheetread.cell_value(r,7)
	face_plus_plus_count = male_count + female_count
	source = ""
	final_gender = ""
	if((male_count > 0 or female_count > 0) and face_plus_plus_count >= 4):
		confidence = float(male_count)/float(male_count + female_count)
		# flip the confidence if it is actually a female 
		if(confidence < 0.5):
			confidence = 1 - confidence
		# if face plus plus rendered a confidence of over 80%, keep the face_plus_plus result 
		if(confidence >= 0.8):
			source = javasheetread.cell_value(r,12) 
			final_gender = javasheetread.cell_value(r,11)
			genderize_confidence = javasheetread.cell_value(r,4)
			genderize_gender = javasheetread.cell_value(r,3)
			# check if genderize had a more accurate prediction than face_plus_plus 
			# we do not need to check if genderize confidence is less than 0.75, becuase
			# that was the cut off to test the name with face_plus_plus 
			if(len(genderize_gender) != 0 and genderize_confidence > confidence):
				 confidence = genderize_confidence
				 source = "genderize"
				 final_gender = genderize_gender
		# if neither face_plus_plus nor genderize had a 80% confidence, we take the bert result 
		else:
			source = "custom_bert"
			final_gender = javasheetread.cell_value(r,21)
	else:
		source = "custom_bert"
		final_gender = javasheetread.cell_value(r,21)
	if(final_gender != javasheetread.cell_value(r,21)):
		differences += 1
		flagIndex = "Z" + str(r + 1)
		javasheetwrite[flagIndex] = confidence
	genderIndex = "X" + str(r + 1)
	sourceIndex = "Y" + str(r + 1)
	javasheetwrite[genderIndex] = final_gender
	javasheetwrite[sourceIndex] = source

	if(r%25==0):
		print(str(r))
		xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/cs224u/gender_eval.xlsx')

xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/cs224u/gender_eval.xlsx')
print(differences)