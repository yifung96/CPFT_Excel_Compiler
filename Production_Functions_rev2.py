# -*- coding: utf-8 -*-
"""
Created on Sat Jul 24 11:16:07 2021

@author: USER
"""
import pandas as pd
import matplotlib.pyplot as plt
import statistics as stat
import numpy as np

class extractData(object):
    def __init__(self,testDF,fileParam,checkKw):
        self.testDF = testDF
        self.fileParam = fileParam
        self.checkKw = checkKw
        datDict,topParam = self.run(self.testDF,self.fileParam,self.checkKw)
        self.datDict = datDict
        self.topParam = topParam
        
    def run(self,testDF,fileParam,checkKw):
        ##  Extract Data from DataFrame
        maxCol = testDF.shape[1]
        maxRow = testDF.shape[0]
        #   Select Columns based on USL and LSL
        topParam = {}
        datDict = {}
        keywords = ['Skew','Temp','VDDA','VDDIO']
        Delimiter = 'OS'
        cnt = 0
        for col in range(fileParam[5],maxCol,1):
            if checkKw:
                keyParam = testDF.columns[col]
                for key in keywords:
                    if key in keyParam:
                        cnt+=1
                        topParam[key] = testDF.iloc[fileParam[4]:maxRow-2,col] 
                    
                if cnt == 4 or keyParam == Delimiter:
                    checkKw = False
            else:
                usl = testDF.iloc[fileParam[1],col]
                lsl = testDF.iloc[fileParam[2],col]
                
                if pd.isna(usl):
                    continue
                else:
                    if usl != lsl:
                        ind = testDF.iloc[fileParam[4]:maxRow-2,col].sum()
                        length = testDF.iloc[fileParam[4]:maxRow-2,col].shape[0]
                        ind2 = (ind/length) - testDF.iloc[fileParam[4],col]
                        if ind == 0 or ind2 == 0:
                            continue
                        else:
                            parType = testDF.columns[col]
                            #   Trim parType
                            if "." in parType:
                                parType = (parType.split(".",1))[0]
                            parName = (parType+testDF.iloc[fileParam[0],col]).replace('_','')
                            #   Reduce name length
                            if len(parName)>30:
                                if 'Average' in parName:
                                    parName = parName.replace('Average','Avg')
                                if 'Accuracy' in parName:
                                    parName = parName.replace('Accuracy','Acc')
                                if 'MOTION' in parName:
                                    parName = parName.replace('MOTION','Mot')
                                if 'Default' in parName:
                                    parName = parName.replace('Default','Def')
                                if 'Manual' in parName:
                                    parName = parName.replace('Manual','Man')
                                if 'IDDVDDIO' in parName:
                                    parName = parName.replace('IDDVDDIO','VDDIO')
                                if 'IDDPDVDDIO' in parName:
                                    parName = parName.replace('IDDPDVDDIO','VDDIO')
                                        
                            if parType not in datDict.keys():
                                datDict[parType] = {}
                            if parName not in datDict[parType].keys():
                                datDict[parType][parName] = {}
                            datDict[parType][parName]['Par'] = {}
                            datDict[parType][parName]['Par']['Name'] = parName
                            datDict[parType][parName]['Par']['Usl'] = usl
                            datDict[parType][parName]['Par']['Lsl'] = lsl
                            datDict[parType][parName]['Par']['Unit'] = testDF.iloc[fileParam[3],col]
                            datDict[parType][parName]['Data'] = testDF.iloc[fileParam[4]:maxRow-2,col] 
                    else:
                        continue
        return datDict,topParam        
 
class storeData(object):
    def __init__(self,datDict,topParam,resBook):
        self.datDict = datDict
        self.topParam = topParam
        self.resBook = resBook
        dataDF,bplimit = self.run(self.datDict,self.topParam,self.resBook)
        self.dataDF = dataDF
        self.bplimit = bplimit
    
    def run(datDict,topParam,resBook):
        # Data Sheet
        copyDict = {}
        bplimit = {}
        for top in topParam:
            copyDict[top] = list(topParam[top])
        for dat in datDict:
            for name in datDict[dat]:
                copyDict[name] = list(datDict[dat][name]['Data'])
                bplimit[name] = {}
                bplimit[name]['Usl'] = datDict[dat][name]['Par']['Usl']
                bplimit[name]['Lsl'] = datDict[dat][name]['Par']['Lsl']
                
        dataDF = pd.DataFrame(copyDict)
        dataDF.to_excel(resBook, sheet_name = "Datalog")
        
        return dataDF,bplimit
    
class plots(object):
    def __init__(self,mode,datDict,fileParam):
        self.mode = mode
        self.datDict = datDict
        self.fileParam = fileParam
        if self.mode == 1:
            self.datDict = self.capAnalyse(self.datDict,self.fileParam)
    
    def capAnalyse(self,datDict,fileData):
        ## To calculate (Lsl,Usl,Mean,StdDev Within,StdDev Overall,CPL,CPU,CPK,PPL,PPU,PPK) 
        rnd = 5 #Round Off
        sig = fileData[6]
        for items in datDict:
            for name in datDict[items]:
                dat = datDict[items][name]['Data']
                # Calculations
                usl = datDict[items][name]['Par']['Usl']
                lsl = datDict[items][name]['Par']['Lsl']
                smpN = len(dat)
                smpMean = stat.mean(dat)
                stDevOV = stat.pstdev(dat)
                stDevWT = stat.stdev(dat)
                cpu = (usl-smpMean)/(sig*stDevWT)
                cpl = (smpMean - lsl)/(sig*stDevWT)
                cpk = np.min([cpl,cpu])
                ppu = (usl-smpMean)/(sig*stDevOV)
                ppl = (smpMean - lsl)/(sig*stDevOV)
                ppk = np.min([ppl,ppu]) 
                # Store in Dict
                datDict[items][name]['Par']['Smp_N'] = smpN
                datDict[items][name]['Par']['Smp_Mean'] = round(smpMean,rnd)
                datDict[items][name]['Par']['StDev_Overall'] = round(stDevOV,rnd)
                datDict[items][name]['Par']['StDev_Within'] = round(stDevWT,rnd)
                datDict[items][name]['Par']['Cpu'] = round(cpu,rnd)
                datDict[items][name]['Par']['Cpl'] = round(cpl,rnd)
                datDict[items][name]['Par']['Cpk'] = round(cpk,rnd)
                datDict[items][name]['Par']['Ppu'] = round(ppu,rnd)
                datDict[items][name]['Par']['Ppl'] = round(ppl,rnd)
                datDict[items][name]['Par']['Ppk'] = round(ppk,rnd)
            ##  Create Summary Table in Excel
        datatoDF = []
        #   Convert Data to DataFrame
        for items in datDict:
            for name in datDict[items]:
                datatoDF.append(datDict[items][name]['Par'])
        resDF = pd.DataFrame(datatoDF)
        resDF.to_excel(resBook, sheet_name = "CASummary")
        
        return datDict