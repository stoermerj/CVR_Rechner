#Analysis of Unique Page Views Online Antrag
import numpy as np
import pandas as pd
import glob

#empty global arrays for files and dataframes
output = []
filesDf = []

#find all excel files from GA in directory to analyze
def fileAnalyzer():
    path = r"C:\Users\stoermerj\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Programme\AnalyticsData"
    files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
    for f in files:
        df = pd.read_excel(f, sheet_name="Dataset1")
        filesDf.append(df)

#Analyze each Excel and clean it from unecessary data
def loopFunction():
    for df in filesDf:
        #remove non ecommerce sites
        dfOA = df[df["Page"].str.contains(r"vpv")==True].reset_index(drop=True)

        #remove save
        dfOA1 = dfOA[dfOA ["Page"].str.contains(r"save|direct")==False]

        #remove non funnel sites
        dfOA2 = dfOA1[dfOA1 ["Page"].str.contains(r"0")==True]

        #replace new URLs with old structure to combine data
        dfOA2["Page"] = dfOA2["Page"].str.replace("online-antrag/", "")
        dfOA2["Page"] = dfOA2["Page"].str.replace("online-antrag.hiscox.de", "")
        dfOA2["Page"] = dfOA2["Page"].str.replace("makler.hiscox.de", "")

        #sort values by page name
        dfOA3 = dfOA2.sort_values(by="Page").reset_index(drop=True)

        #group each step
        dfOA4 = dfOA3.groupby("Page").sum().reset_index()

        #remove added percentage columns
        dfOA5 = dfOA4.loc[:,("Page", "Page Views", "Unique Page Views", "Entrances")]

        #calculate bounce rate and exit rate
        dfOA5["% Double Page View"] = dfOA5["Unique Page Views"] / dfOA5["Page Views"]

        #Finale CVRs
        CVR1 = dfOA5.loc[1,"Unique Page Views"] / dfOA5.loc[0,"Unique Page Views"]  
        CVR2 = dfOA5.loc[2,"Unique Page Views"] / dfOA5.loc[1,"Unique Page Views"] 
        CVR3 = dfOA5.loc[3,"Unique Page Views"] / dfOA5.loc[2,"Unique Page Views"] 
        CVR4 = dfOA5.loc[4,"Unique Page Views"] / dfOA5.loc[3,"Unique Page Views"] 
        CVR5 = dfOA5.loc[5,"Unique Page Views"] / dfOA5.loc[4,"Unique Page Views"]
        finalCVR = dfOA5.loc[5,"Page Views"] / dfOA5.loc[0,"Unique Page Views"] 

        #CVR hinzufügen
        cvr = [CVR1, CVR2, CVR3, CVR4, CVR5, finalCVR]
        dfOA5["CVR"] = pd.Series(cvr)

        #print everything
        print(dfOA5)

        #append to global 
        output.append(dfOA5)

#export data to excel file
def writeToExcel():
    
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('outputNew2.xlsx', engine='xlsxwriter')

    y = 0
    for x in output:
        # Convert the dataframe to an XlsxWriter Excel object.
        x.to_excel(writer, sheet_name=str(y), index=False)
        y = y + 1

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

fileAnalyzer()
loopFunction()
writeToExcel()
