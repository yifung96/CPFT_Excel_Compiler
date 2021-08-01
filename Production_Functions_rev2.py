# -*- coding: utf-8 -*-
"""
Created on Sat Jul 24 11:16:07 2021

@author: YF
"""

import pandas as pd
import numpy as np
import statistics as stat
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

class extractData(object):
    def __init__(self,testDF,fileParam,checkKw):
        self.testDF = testDF
        self.fileParam = fileParam
        self.checkKw = checkKw
        self.run(self.testDF,self.fileParam,self.checkKw)
        
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
        self.dataDict = datDict
        self.topParam = topParam
        
class storeData(object):
    def __init__(self,datDict,topParam,resBook):
        self.datDict = datDict
        self.topParam = topParam
        self.resBook = resBook
        self.run(self.datDict,self.topParam,self.resBook)
    
    def run(self,datDict,topParam,resBook):
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
        self.dataDF = dataDF
        self.bplimit = bplimit
        
class plots(object):
    def __init__(self,mode,dataDict,fileParam,resBook,resPath,dataDF,bplimit,mtb):
        self.resBook = resBook
        self.mode = mode
        self.datDict = dataDict
        self.fileParam = fileParam
        self.resPath = resPath
        self.mtb = mtb
        if self.mode == 1:
            self.datDict = self.capAnalyse(self.datDict,self.fileParam,self.resBook)
            if self.mtb:
                self.plotCA2(self.datDict,self.resBook)
            else:
                self.plotCA(self.datDict,self.resBook,self.resPath)
        else:
            self.dataDF = dataDF
            self.bplimit = bplimit
            self.plotBoxplot(self.dataDF,self.bplimit,self.resBook,self.resPath)
    
    def capAnalyse(self,datDict,fileData,resBook):
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
    
    def plotCA(self,datDict,resBook,resPath):
        # Plot Capability Graph
        i = 2
        j = 2
        worksheet = resBook.sheets["CASummary"]
        for items in datDict:
            for name in datDict[items]:
                imgcell = "R"+str(i)
                linkcell = "A"+str(j)
                urlcell = "U"+str(i+5)
                img = resPath+'\\CP_Plots'+str(j)+'.png'
                #worksheet = resBook.sheets["CASummary"]
                # Generate probability density function 
                dataset = list(datDict[items][name]['Data'].values)
                Lsl = datDict[items][name]['Par']['Lsl']
                Usl = datDict[items][name]['Par']['Usl']
                unit = datDict[items][name]['Par']['Unit']
                stDev = datDict[items][name]['Par']['StDev_Overall']
                # Methods to find bin size [Scotts, Freedman, Sturge]
                calc_bin = self.scotts(dataset,stDev)
                #calc_bin = self.freedman_diaconis(Lsl,Usl,dataset)
                #calc_bin = self.sturge(dataset)
                print(name,"",calc_bin)
                if calc_bin == 0:
                    continue
                else:
                    # Plot Histogram
                    plt.hist(dataset, bins=calc_bin, color="lightgrey", edgecolor="grey", density=True)
                    sns.kdeplot(dataset, color='black',label="Density")
                    plt.axvline(Lsl, linestyle="--", color="red", label="LSL")
                    plt.axvline(Usl, linestyle="--", color="orange", label="USL")
                    plt.title(name)
                    plt.xlabel("")
                    plt.ylabel(unit)    
                    plt.legend()
                    plt.savefig(img)
                    #plt.show()
                    worksheet.insert_image(imgcell,img)
                    # Insert Link
                    worksheet.write_url(linkcell,"internal:'CASummary'!"+urlcell,string="Link")
                    i+=20
                    j+=1
                    plt.close()
        resBook.save()
        
    def freedman_diaconis(self,Lsl,Usl,dataset):
        ## Often will have too many Bins
        N = len(dataset)
        IQR  = stats.iqr(dataset, rng=(25, 75), scale=1, nan_policy="omit")
        bw = (2*IQR)/np.power(N,1/3)
        print(max(dataset),min(dataset),IQR,bw)
        if bw == 0:
            calc_bin = 10
        else:
            calc_bin = int((max(dataset)-min(dataset))/bw)
        return calc_bin

    def sturge(self,dataset):
        ## Good for N<30
        N = len(dataset)
        calc_bin = int(np.ceil(np.log2(N) + 1))
        return calc_bin

    def scotts(self,dataset,stDev):
        N = len(dataset)
        bw = (3.49*stDev)/np.power(N,1/3)
        calc_bin = int((max(dataset)-min(dataset))/bw)
        return calc_bin
    
    def plotBoxplot(self,dataDF,bplimit,resBook,resPath):
        tmpDF = (pd.DataFrame(bplimit)).transpose()
        tmpDF.to_excel(resBook, sheet_name = "BPSummary")
        #   Plot Boxplot
        startplot_ind = False
        startplot_kw = 'VDDIO'
        k = 2
        l = 2
        worksheet = resBook.sheets["BPSummary"]
        for col in dataDF:
            if col == startplot_kw:
                startplot_ind = True
            elif startplot_ind:
                #   Declare parameters
                vddacell = "G"+str(k)
                vddiocell = "P"+str(k)
                linkcell = "D"+str(l)
                urlcell = "O"+str(k+10)
                vddaimg = resPath+'\\VDDA_Plots'+str(l)+'.png'
                vddioimg = resPath+'\\VDDIO_Plots'+str(l)+'.png'
                
                print(col)
                usl = bplimit[col]['Usl']
                lsl = bplimit[col]['Lsl']
                #   Process VDDA Boxplot
                vdda_bp = dataDF.boxplot(column=col,by=['Skew','Temp','VDDA'],grid=False, rot=90, fontsize=8)
                vdda_bp.axhline(usl, linestyle="--", color="black", label="Usl")
                vdda_bp.axhline(lsl, linestyle="--", color="black", label="Lsl")
                plt.suptitle("")
                plt.savefig(vddaimg,bbox_inches='tight')
                worksheet.insert_image(vddacell,vddaimg)
                #   Process VDDIO Boxplot
                vddio_bp = dataDF.boxplot(column=col,by=['Skew','Temp','VDDIO'],grid=False, rot=90, fontsize=8)
                vddio_bp.axhline(usl, linestyle="--", color="black", label="Usl")
                vddio_bp.axhline(lsl, linestyle="--", color="black", label="Lsl")
                plt.suptitle("")
                plt.savefig(vddioimg,bbox_inches='tight')
                worksheet.insert_image(vddiocell,vddioimg)
                #   Insert Links
                worksheet.write_url(linkcell,"internal:'BPSummary'!"+urlcell,string="Link")
                k+=22
                l+=1
            else:
                continue
        resBook.save()
        
    def plotCA2(self,datDict,resBook):
        # Plot Capability Graph
        i = 2
        j = 2
        worksheet = resBook.sheets["CASummary"]
        for items in datDict:
            for name in datDict[items]:
                linkcell = "A"+str(j)
                urlcell = "U"+str(i+5)
                # Insert Link
                worksheet.write_url(linkcell,"internal:'CASummary'!"+urlcell,string="Link")
                i+=20
                j+=1
        resBook.save()
        
class mtb_cmd(object):
    def __init__(self,dataDict,mode,resPath):
        self.mode = mode
        self.dataDict = dataDict
        self.resPath = resPath
        if self.mode == 1:
            self.mtb_cmd1(self.dataDict,self.resPath)
        else:
            self.mtb_cmd2(self.dataDict,self.resPath)
            self.mtb_cmd3(self.dataDict,self.resPath)
            
    def mtb_cmd1(self,dataDict,resPath):
        #   Write commands to textfile
        lines = []
        for items in dataDict:
            for name in dataDict[items]:
                #   Get required data
                Lsl = dataDict[items][name]['Par']['Lsl']
                Usl = dataDict[items][name]['Par']['Usl']
                N = len(list(dataDict[items][name]['Data'].values))
                #   Write into excel
                lines.append("MTB > Capa " +"'"+name+"' "+str(N)+";")
                lines.append("SUBC> Lspec "+str(Lsl)+";")
                lines.append("SUBC> Uspec "+str(Usl)+";")
                lines.append("SUBC> Pooled;")
                lines.append("SUBC> AMR;")
                lines.append("SUBC> UnBiased;")
                lines.append("SUBC> OBiased;")
                lines.append("SUBC> Toler 6;")
                lines.append("SUBC> Within;")
                lines.append("SUBC> Overall;")
                lines.append("SUBC> CStat;")
                lines.append("SUBC> GSAVE "+'"'+name+'"'+";")
                lines.append("SUBC> PNGH.")
        cmd = open(resPath+'\\cmd.txt',"w")
        for line in lines:
            cmd.write(line+"\n")
        cmd.close()   

    def mtb_cmd2(self,dataDict,resPath):
        #   Write commands to textfile
        lines = []
        for items in dataDict:
            for name in dataDict[items]:
                #   Get required data
                Lsl = dataDict[items][name]['Par']['Lsl']
                Usl = dataDict[items][name]['Par']['Usl']
                #N = len(list(datDict[items][name]['Data'].values))
                #   Write into excel
                lines.append("MTB > Boxplot " +"('"+name+"') * Skew"+";")
                lines.append("SUBC> Group Temp VDDIO"+";")
                lines.append("SUBC> Overlay"+";")
                lines.append("SUBC> Reference 2 "+str(Usl)+" "+str(Lsl)+";")
                lines.append("SUBC> IQRBox;")
                lines.append("SUBC> Outlier.")
        cmd = open(resPath+'\\VDDIO_cmd.txt',"w")
        for line in lines:
            cmd.write(line+"\n")
        cmd.close()
        
    def mtb_cmd3(self,dataDict,resPath):
        #   Write commands to textfile
        lines = []
        for items in dataDict:
            for name in dataDict[items]:
                #   Get required data
                Lsl = dataDict[items][name]['Par']['Lsl']
                Usl = dataDict[items][name]['Par']['Usl']
                #N = len(list(datDict[items][name]['Data'].values))
                #   Write into excel
                lines.append("MTB > Boxplot " +"('"+name+"') * Skew"+";")
                lines.append("SUBC> Group Temp VDDA"+";")
                lines.append("SUBC> Overlay"+";")
                lines.append("SUBC> Reference 2 "+str(Usl)+" "+str(Lsl)+";")
                lines.append("SUBC> IQRBox;")
                lines.append("SUBC> Outlier.")
        cmd = open(resPath+'\\VDDA_cmd.txt',"w")
        for line in lines:
            cmd.write(line+"\n")
        cmd.close()        
        