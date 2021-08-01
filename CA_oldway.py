# -*- coding: utf-8 -*-
"""
Created on Mon Jul 12 11:16:09 2021

@author: YF
"""

##  Copy from folder into excel
import os
import xlsxwriter

imgPath = 'ResultFolder\\Prod_Data\\Plots'
imgfile = []
with os.scandir(imgPath) as entries:
            for entry in entries:
                ext = os.path.splitext(entry)[1]
                if ext == '.PNG':
                    #print(ext)
                    imgfile.append(entry.name) 
#print(imgfile)

workbook = xlsxwriter.Workbook('Plots.xlsx')
worksheet = workbook.add_worksheet('Plots')

i = 2
j = 2
for img in imgfile:
    imgcell = "R"+str(i)
    linkcell = "A"+str(j)
    urlcell = "AE"+str(i+18)
    
    imgloc = imgPath+"\\"+img
    worksheet.insert_image(imgcell,imgloc)
    worksheet.write_url(linkcell,"internal:'Plots'!"+urlcell,string="Link")
    i+=33
    j+=1

workbook.close()