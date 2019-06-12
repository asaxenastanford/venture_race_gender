import xlrd
import openpyxl
import requests
import os, shutil

def run(command):
    print(command)
    exit_status = os.system(command)
    if exit_status > 0:
        exit(1)

javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/face_plus_plus/names_to_run.xlsx'
wbJ = xlrd.open_workbook(javaFile)
javasheetread = wbJ.sheet_by_index(0)

for r in range(2, 1668):
	firstName = javasheetread.cell_value(r,0)
	lastName = javasheetread.cell_value(r,1)
	name = firstName + lastName 
	name = name.replace(" ", "")
	name = name.replace("\'", "")
	print("Name: " + name)
	command = 'python image_scrape_google.py --search ' + name + ' --num_images 5' 
	run(command)


 