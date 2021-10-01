import os

import numpy as np
import openpyxl
import pandas as pd
import os.path
import shutil
import logging
import re
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Border, Side
# from Utils import Utils
from os import path
from Utils import Utils


def folderStructureCreation(dirPath):
    print("Checking Folder structure...", end="")

    folderList = ['in', 'out', 'log', 'error', 'template']

    for folder in folderList:
        folderPath = os.path.join(dirPath, folder)

        # Check if folder exist
        if not path.exists(folderPath):
            os.makedirs(folderPath)

    print("Done!")

    pass


def dropColumnGI(df, var):
    return df.drop(columns=[var,
                            'Reservation Control Number',
                            'How do you plan to use the Moderna vaccine you ordered?',
                            'Select the location of the Moderna vaccinee.'])


def renameColumn(df, param):
    arrPUR = {'Reservation Control Number' + param: 'Reservation Control Number',
              'How do you plan to use the Moderna vaccine you ordered?' + param: 'How do you plan to use the Moderna vaccine you ordered?',
              'Select the location of the Moderna vaccinee.' + param: 'Select the location of the Moderna vaccinee.'}
    return df.rename(columns=arrPUR, inplace=True)


def checkCtrlNumber(controlNumber):
    special_characters = "!@#$%^&*()-+?=,<>/"

    if any(c in special_characters for c in controlNumber):
        return True
    else:
        return False
    pass


def validateCtrlNumber(x, compCode):
    util = Utils()
    compCode = util.companyNameLookUpMethod(compCode)

    regex = "\\b" + compCode + "_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$"
    regexPalex = "\\bPALEX_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$"
    x = x.strip()

    if x == 'None':
        return 'None'
    else:
        if compCode == 'PAL':
            if re.match(regex, x):
                return True
            elif re.match(regexPalex, x):
                return True
            else:
                return False
        else:
            if re.match(regex, x):
                return True
            else:
                return False


def pur1Checker(df):
    df = df.loc[df['pur1'] == 'Yes']

    df = dropColumnGI(df, 'pur1')
    df = df.drop(columns=['pur2',
                          'Reservation Control Number3',
                          'How do you plan to use the Moderna vaccine you ordered?3',
                          'Select the location of the Moderna vaccinee.3',
                          'pur3',
                          'Reservation Control Number4',
                          'How do you plan to use the Moderna vaccine you ordered?4',
                          'Select the location of the Moderna vaccinee.4',
                          'pur4',
                          'Reservation Control Number5',
                          'How do you plan to use the Moderna vaccine you ordered?5',
                          'Select the location of the Moderna vaccinee.5'])

    df['Is SpecialChar Ctrl inValid'] = df.apply(lambda x: checkCtrlNumber(x['Reservation Control Number2']), axis=1)
    df['validateCtrl Format'] = df.apply(lambda x: validateCtrlNumber(x['Reservation Control Number2'], x['Company']),
                                         axis=1)
    renameColumn(df, '2')

    return df


def pur2Checker(df):
    df = df.loc[df['pur2'] == 'Yes']
    df = dropColumnGI(df, 'pur2')
    df = df.drop(columns=['pur1',
                          'Reservation Control Number2',
                          'How do you plan to use the Moderna vaccine you ordered?2',
                          'Select the location of the Moderna vaccinee.2',
                          'pur3',
                          'Reservation Control Number4',
                          'How do you plan to use the Moderna vaccine you ordered?4',
                          'Select the location of the Moderna vaccinee.4',
                          'pur4',
                          'Reservation Control Number5',
                          'How do you plan to use the Moderna vaccine you ordered?5',
                          'Select the location of the Moderna vaccinee.5'])

    df['Is SpecialChar Ctrl inValid'] = df.apply(lambda x: checkCtrlNumber(x['Reservation Control Number3']), axis=1)
    df['validateCtrl Format'] = df.apply(lambda x: validateCtrlNumber(x['Reservation Control Number3'], x['Company']),
                                         axis=1)

    renameColumn(df, '3')
    return df


def pur3Checker(df):
    df = df.loc[df['pur3'] == 'Yes']
    df = dropColumnGI(df, 'pur3')
    df = df.drop(columns=['pur1',
                          'Reservation Control Number2',
                          'How do you plan to use the Moderna vaccine you ordered?2',
                          'Select the location of the Moderna vaccinee.2',
                          'pur2',
                          'Reservation Control Number3',
                          'How do you plan to use the Moderna vaccine you ordered?3',
                          'Select the location of the Moderna vaccinee.3',
                          'pur4',
                          'Reservation Control Number5',
                          'How do you plan to use the Moderna vaccine you ordered?5',
                          'Select the location of the Moderna vaccinee.5'])

    df['Is SpecialChar Ctrl inValid'] = df.apply(lambda x: checkCtrlNumber(x['Reservation Control Number4']), axis=1)
    df['validateCtrl Format'] = df.apply(lambda x: validateCtrlNumber(x['Reservation Control Number4'], x['Company']),
                                         axis=1)

    renameColumn(df, '4')
    return df


def pur4Checker(df):
    df = df.loc[df['pur4'] == 'Yes']
    df = dropColumnGI(df, 'pur4')
    df = df.drop(columns=['pur1',
                          'Reservation Control Number2',
                          'How do you plan to use the Moderna vaccine you ordered?2',
                          'Select the location of the Moderna vaccinee.2',
                          'pur2',
                          'Reservation Control Number3',
                          'How do you plan to use the Moderna vaccine you ordered?3',
                          'Select the location of the Moderna vaccinee.3',
                          'pur3',
                          'Reservation Control Number4',
                          'How do you plan to use the Moderna vaccine you ordered?4',
                          'Select the location of the Moderna vaccinee.4'])

    df['Is SpecialChar Ctrl inValid'] = df.apply(lambda x: checkCtrlNumber(x['Reservation Control Number5']), axis=1)
    df['validateCtrl Format'] = df.apply(lambda x: validateCtrlNumber(x['Reservation Control Number5'], x['Company']),
                                         axis=1)

    renameColumn(df, '5')
    return df


def getData(fileName):
    filePath = os.path.join(inPath, fileName)
    df = pd.read_excel(filePath, dtype=str, na_filter=False)

    # findIncorec(df)
    arrdfFrames = []

    arrPUR = {'Do you have another pending unfulfilled Reservation Control Number or pending Moderna order?': 'pur1',
              'Do you have another pending unfulfilled Reservation Control Number or pending Moderna order?2': 'pur2',
              'Do you have another pending unfulfilled Reservation Control Number or pending Moderna order?3': 'pur3',
              'Do you have another pending unfulfilled Reservation Control Number or pending Moderna order?4': 'pur4'}
    df.rename(columns=arrPUR, inplace=True)

    # New column filename
    fileName = os.path.splitext(fileName)[0]
    companyName = fileName.split("(")[1].replace(")", "")
    df['Company'] = companyName

    df['Is SpecialChar Ctrl inValid'] = df.apply(lambda x: checkCtrlNumber(x['Reservation Control Number']), axis=1)
    df['validateCtrl Format'] = df.apply(lambda x: validateCtrlNumber(x['Reservation Control Number'], x['Company']),
                                         axis=1)

    # Drop Column A-F
    df.drop(df.iloc[:, 1:6], inplace=True, axis=1)

    originalDF = df.drop(df.iloc[:, 10:26], axis=1)

    arrdfFrames.append(originalDF)
    arrdfFrames.append(pur1Checker(df))
    arrdfFrames.append(pur2Checker(df))
    arrdfFrames.append(pur3Checker(df))
    arrdfFrames.append(pur4Checker(df))

    # Merge all df
    df_master = pd.concat(arrdfFrames)

    # New Column concat name for sorting
    df_master['concatFullName'] = df_master['Last Name'].str.lower() + \
                                  df_master['First Name'].str.lower() + \
                                  df_master['Middle Name'].str.lower()

    # sort by name
    df_master.sort_values('concatFullName', inplace=True, ascending=False)

    # drop concatFullName column
    df_master.drop(columns=['concatFullName'], inplace=True)
    # df_master.drop(columns=['ID'], inplace=True)

    # Convert birthdate
    df_master['Birthdate'] = pd.to_datetime(df_master['Birthdate'])  # convert the column to a datetime object
    df_master['Birthdate'] = df_master['Birthdate'].dt.strftime('%m/%d/%Y')  # format the object

    # Check Duplicate Control Number
    df_master['Is Ctrlnum Dup'] = df_master.duplicated(subset="Reservation Control Number", keep=False)

    return df_master


def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
    companyDir = out + "/"
    srcFile = companyDir + outputFilename + ".xlsx"

    if not os.path.isfile(srcFile):
        shutil.copy(tempLTGC_Path, srcFile)

    return companyDir + outputFilename + ".xlsx"


def generateErrorCtrlNumberFormat(df):
    df_specialChar = df.loc[(df['Is SpecialChar Ctrl inValid'] == True) & (df['validateCtrl Format'] == False)]

    errMsg = []

    groups = df_specialChar.groupby('Company')
    for comp, records in groups:
        for j, row in records.iterrows():
            errMsg.append("Error: ID[ " + str(row['ID']) + " ] - " + row['Reservation Control Number'] +
                          " Wrong Control Number Format")

            generateErrorLog(errMsg, row['Company'], 'withSpecialChar')

    pass


def generateErrorCtrlNumberDup(df):
    df_ctrldup = df.loc[df['Is Ctrlnum Dup'] == True]

    arrid = []
    errMsg = []

    groupsctrl = df_ctrldup.groupby('Reservation Control Number')
    for comp, records in groupsctrl:
        # print(comp)
        for j, row in records.iterrows():
            arrid.append(row['ID'])

        errMsg.append("Error: " + comp + " - Duplicate Control Number found in ID [" + ','.join(arrid) + "]")

        arrid.clear()

        generateErrorLog(errMsg, row['Company'], 'DuplicateControlNumber')

    # print(errMsg)
    pass


def generateErrorLog(errMsg, companyCode, arg):
    util = Utils()
    if len(errMsg):
        util.createSubCompanyFolder(companyCode, errPath)
        f = open(
            errPath + "/" + companyCode + "/" + companyCode + "_" + arg + "_err_log_" + dateTime + ".txt",
            "a")
        for err in errMsg:
            f.writelines(err + "\n")

        errMsg.clear()
    pass


def dropInvalidControlNumber(df):
    df.drop(df[df['Is SpecialChar Ctrl inValid'] == True].index, inplace=True)
    df.drop(df[df['validateCtrl Format'] == False].index, inplace=True)

    pass


def dropColumn3Last(df):
    df.drop(columns=['Is SpecialChar Ctrl inValid',
                     'validateCtrl Format'], inplace=True)

    df.drop(columns=['Is Ctrlnum Dup'], inplace=True)

    pass


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m%d%y%H%M%S")

    #MAC OS Path
    dirPath = r"/Users/Ran/Documents/Vaccine/RegOfModernaOrderForHH"

    inPath = os.path.join(dirPath, "in")
    outPath = os.path.join(dirPath, "out")
    errPath = os.path.join(dirPath, "error")
    logPath = os.path.join(dirPath, "log")
    backupPath = os.path.join(dirPath, "backup")
    templateFilePath = os.path.join(dirPath, "template/template_v1.xlsx")

    outFilename = 'Registration_of_Moderna_Order_for_HH_Conso_' + dateTime

    # Folder Structure Creation
    folderStructureCreation(dirPath)

    logging.info("==============================================================")
    logging.info("Running Scpirt: REgistration of Moderna Order for HH Consolidation......")
    logging.info("==============================================================")

    # Get all Files in
    arrFilenames = os.listdir(inPath)
    arrdf = []

    for inFile in arrFilenames:
        if not inFile == ".DS_Store":
            print("Reading: " + inFile + "......")

            arrdf.append(getData(inFile))

    master_df = pd.concat(arrdf)

    generateErrorCtrlNumberFormat(master_df)
    generateErrorCtrlNumberDup(master_df)

    dropInvalidControlNumber(master_df)

    dropColumn3Last(master_df)

    master_df.drop_duplicates(subset=['Reservation Control Number'], keep=False, inplace=True)

    # Create copy of template file and save it to out folder
    templateFile = duplicateTemplateLTGC(templateFilePath, outPath, outFilename)

    # Write df_master(consolidated/append data) to excel
    writer = pd.ExcelWriter(templateFile, engine='openpyxl', mode='a')
    writer.book = openpyxl.load_workbook(templateFile)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    master_df.to_excel(writer, sheet_name="Form1", startrow=1, header=False, index=False)
    writer.save()
