from numpy import unicode_
import xlwt
import os
import openpyxl 
import sys
import pdftotext
import zipfile
import shutil

from AddNumber import addNumber
from Rename9DW import rename9DW
from ExtractKeyWords import extractKeyWords
from RemoveZero import removeZero
from ChineseDel import chineseDel
from Pdf_combine import pdf_combine
from RemoveNumber import removeNumber
from Rename7DX import rename7DX
import time
import logging

logging.basicConfig(level=logging.DEBUG)
time1 = time.time()

path="C:\\Users\\nwcdi\\Downloads\\archives\\test\\155"
file="/mnt/c/Project/workspace/data/pdf/91f18a7c56f3426ea937e7001289b5b6.zip"
file="/mnt/c/Project/workspace/data/pdf/1000.zip"

file=sys.argv[1]

print(sys.argv)

#    zip_ref.extractall(dest)

fileformat=os.path.splitext(file)[-1][1:]

# the folder that zip unzip location
zipFolder=os.path.splitext(file)[0]

PyChoose=int(sys.argv[2])

offset=sys.argv[3] #if len(sys.argv)>=4 else "0"



if PyChoose==0:
    target=os.path.dirname(file)+os.sep+offset
    unzipFolder=zipFolder+os.sep+offset

    #unzip file
    shutil.unpack_archive(file,unzipFolder,fileformat)    
    addNumber(unzipFolder,offset)
    extractKeyWords(unzipFolder)
    removeZero(unzipFolder)
    pdf_combine(unzipFolder)
else:
    
    target=os.path.dirname(zipFolder)+os.sep+offset  

    unzipFolder=zipFolder
    print(target,unzipFolder)

    shutil.unpack_archive(file,unzipFolder,fileformat) 

    if PyChoose==1:
        chineseDel(unzipFolder)
    elif PyChoose==2:
        rename7DX(unzipFolder,offset)
    elif PyChoose==3:
        rename9DW(unzipFolder,offset)


#The target+fileformat is the target zip file
shutil.make_archive(target,fileformat,zipFolder)

#remove temp folder
shutil.rmtree(zipFolder)
time2 = time.time()
print('总共耗时：%s s.' %(time2 - time1))
