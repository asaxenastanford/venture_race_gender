import uuid
import sys
import collections
import argparse
import json
import xlrd
import openpyxl
import itertools
import logging
import re
import os
from urllib.request import urlopen, Request

from bs4 import BeautifulSoup
from facepplib import FacePP


def configure_logging():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    handler = logging.StreamHandler()
    logger.addHandler(handler)
    return logger

logger = configure_logging()

REQUEST_HEADER = {
    'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.134 Safari/537.36"}

def soup_call(url, header):
    response = urlopen(Request(url, headers=header))
    return BeautifulSoup(response, 'html.parser')

def get_url(q):
    return "https://www.google.co.in/search?q=%s&source=lnms&tbm=isch" % q

def images_extraction_soup(soup):
    image_elements = soup.find_all("div", {"class": "rg_meta"})
    metadata_dicts = (json.loads(e.text) for e in image_elements)
    link_type_records = ((d["ou"], d["ity"]) for d in metadata_dicts)
    return link_type_records

def images_extraction(q, num_images):
    url = get_url(q)
    soup = soup_call(url, REQUEST_HEADER)
    type_r = images_extraction_soup(soup)
    return itertools.islice(type_r, num_images)

def updateSheets(genderDic,ethnicityDic):
    javaFile = '/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/face_plus_plus/names_to_run.xlsx'
    wbJ = xlrd.open_workbook(javaFile)
    javasheetread = wbJ.sheet_by_index(0)

    xfile1 = openpyxl.load_workbook('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/face_plus_plus/names_to_run.xlsx')
    javasheetwrite = xfile1['Sheet1'] 
    last_row = int(javasheetread.cell_value(0,5))

    male_count = genderDic["Male"]
    female_count = genderDic["Female"]

    maleInd = "D" + str(last_row + 1)
    femInd = "E" + str(last_row + 1)
    finInd = "F" + str(last_row + 1)
    countInd = "G" + str(last_row + 1)
    ethInd = "H" + str(last_row + 1)
    nameId = "I" + str(last_row + 1)
    javasheetwrite["F1"] = (last_row + 1)

    javasheetwrite[maleInd] = male_count
    javasheetwrite[femInd] = female_count

    if(male_count > 0 or female_count >0):
        male_confidence =  float(male_count)/float(male_count + female_count)
        if(male_confidence > 0.5):
            javasheetwrite[finInd] = "male"
            # print("male classification")
        else:
            javasheetwrite[finInd] = "female"
            # print("female classification")

    actualEth = ethnicityDic.most_common(1) # the most commonly found ethnicity among the images sampled 
    if(len(actualEth) > 0):
        javasheetwrite[countInd] = actualEth[0][1]
        javasheetwrite[ethInd] = (actualEth[0][0].lower())


    xfile1.save('/Users/aakankshasaxena/Documents/Sophomore/Strebulaev/face_plus_plus/names_to_run.xlsx')

def find_images(images, num_images): #save_directory, 
    facepp = FacePP(api_key='prYP6L8GFA06ycwOYdRGpOowEOLWrhr5', api_secret='aCrsKLu_tnC-3srCI4B2ID6NrXhV1WA_',url='https://api-us.faceplusplus.com')
    genderDic = {}
    ethnicityDic = collections.Counter()
    genderDic["Male"] = 0
    genderDic["Female"] = 0
    print("analyzing pictures")

    for i, (url, image_type) in enumerate(images):
        print("Next image!")
        print("Url: " + url)
        try:
            image = facepp.image.get(image_url=url, return_attributes=['gender'])
            print("True value: " + image.faces[0].gender['value'])
            if(image.faces[0].gender['value'] == "Male"):
                print("Stored male")
                genderDic["Male"] += 1
            elif(image.faces[0].gender['value'] == "Female"):
                print("Stored female")
                genderDic["Female"] += 1
            ethnicityDic[image.faces[0].ethnicity['value']] += 1

        except Exception as e:
            logger.exception(e)

    updateSheets(genderDic, ethnicityDic)

def run_program(query, num_images=100): 
    q = '+'.join(query.split())
    img = images_extraction(q, num_images)
    find_images(img, num_images) 

def main():
    facepp = FacePP(api_key='prYP6L8GFA06ycwOYdRGpOowEOLWrhr5', api_secret='aCrsKLu_tnC-3srCI4B2ID6NrXhV1WA_',url='https://api-us.faceplusplus.com')
    parser = argparse.ArgumentParser()
    parser.add_argument('-s', '--search', default='JohnSmith', type=str)
    parser.add_argument('-n', '--num_images', default=1, type=int)
    args = parser.parse_args()
    run_program(args.search, args.num_images) 

if __name__ == '__main__':
    main()