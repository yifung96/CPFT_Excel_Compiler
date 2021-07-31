# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 20:38:46 2021

@author: USER
"""
import os
import time
import pandas as pd
import Production_Functions_rev2 as func

##  Create Folder to store database and result
if not os.path.exists('DataFolder\\Prod_Data'):
    os.makedirs('DataFolder\\Prod_Data')
if not os.path.exists('ResultFolder\\Prod_Data'):
    os.makedirs('ResultFolder\\Prod_Data')

#################################   Change Setting here   #############################################
filename = 'testfile.xlsx'
sheetname = 'FF' 
#   [1: CP Analysis|| 2: BoxPlot]
mode = 1
#######################################################################################################

##  Initiate Time
t0 = time.time() 
##  Fixed Setting
fileParam = [0,2,3,4,5,1,3]  #fileData = [ParName, Usl, Lsl, Unit, startrow, startcol, sigma]
if mode == 2:
    checkKw = True
else:
    checkKw = False
##  Input Filepath
filePath = 'DataFolder\\Prod_Data\\'+filename
print("TestFile: ", filePath)
##  Output Folder
currTime = (time.strftime("%m:%d:%H:%M:%S",time.localtime())).replace(':','')   
resPath = 'ResultFolder\\Prod_Data\\'+currTime+'mode'+str(mode)
os.makedirs(resPath)
resFile = resPath+'\\Summary.xlsx'

##  Process Datalog
#   Convert Excel sheet to Dataframe
testDF = pd.read_excel(filePath,sheet_name = sheetname)
print("Datalog Dimension: ",testDF.shape)
datDict,topParam = func.extractData(testDF,fileParam,checkKw)
resBook = pd.ExcelWriter(resFile, engine = 'xlsxwriter')
dataDF,bplimit = func.storeData(datDict,topParam,resBook)
print("Complete Compilation -> Processing Plots")



