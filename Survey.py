import boto3
import os
import glob
import xlwt
import json
import re
import xlrd
import shutil
import pandas as pd
from pandas import ExcelWriter
import configparser
from pathlib import Path

####### Read Configration file ########
config = configparser.ConfigParser()
config_file = './config/default.ini'
config.read(config_file)

######### check folder #############
tempFolder = config['File']['tempFolder']
inputFolder = config['File']['inputFolder']
outputFolder = config['File']['outputFolder']
downloadFolder = config['File']['downloadFolder']

# temp folder
if os.path.isdir(tempFolder):
    print('✅ ' + tempFolder + ' Directory Exists')
    pass
else:
    print('❌ ' + tempFolder + ' Not Exists')
    os.mkdir(tempFolder)
    print('✅ ' + tempFolder + ' Directory Created')

# input folder
if os.path.isdir(inputFolder):
    print('✅ ' + inputFolder + ' Directory Exists')
    pass
else:
    print('❌ ' + downloadFolder + ' Not Exists')
    os.mkdir(inputFolder)
    print('✅ ' + inputFolder + ' Directory Created')

# output folder
if os.path.isdir(outputFolder):
    print('✅ ' + outputFolder + ' Directory Exists')
    pass
else:
    print('❌ ' + downloadFolder + ' Not Exists')
    os.mkdir(outputFolder)
    print('✅ ' + outputFolder + ' Directory Created')

# download folder
if os.path.isdir(downloadFolder):
    print('✅ ' + downloadFolder + ' Directory Exists')
    pass
else:
    print('❌ ' + downloadFolder + ' Not Exists')
    os.mkdir(downloadFolder)
    print('✅ ' + downloadFolder + ' Directory Created')


SurveyAnswer_en_xls = config['File']['outputFolder'] + '/' + config['File']['SurveyAnswers_en'] + '.xls'
SurveyAnswer_zh_xls = config['File']['outputFolder'] + '/' + config['File']['SurveyAnswers_zh'] + '.xls'
SurveyAnswer_fr_xls = config['File']['outputFolder'] + '/' + config['File']['SurveyAnswers_fr'] + '.xls'

#######################################
client = config['AWS']['prefix'].split('/')[0]

####set up excel encoding ##############
book_en = xlwt.Workbook(encoding='utf-8', style_compression=0)
book_zh = xlwt.Workbook(encoding='utf-8', style_compression=0)
book_fr = xlwt.Workbook(encoding='utf-8', style_compression=0)

####### excel spreadsheet ##############
sheet_en = book_en.add_sheet(config['File']['SurveyAnswers_en'], cell_overwrite_ok=True)
sheet_zh = book_zh.add_sheet(config['File']['SurveyAnswers_zh'], cell_overwrite_ok=True)
sheet_fr = book_fr.add_sheet(config['File']['SurveyAnswers_fr'], cell_overwrite_ok=True)

####### add style for header fonts #####
Headerfont = xlwt.Font()  # Create the Font
Headerfont.name = 'Arial'
Headerfont.height = 280
Headerfont.bold = True
HeaderStyle = xlwt.XFStyle()
HeaderStyle.font = Headerfont

####### add style for question fonts #####
Questionfont = xlwt.Font()
Questionfont.name = 'Arial'
Questionfont.height = 240
QuestionStyle = xlwt.XFStyle()
QuestionStyle.font = Questionfont

####### AWS Configuration ###########
s3 = boto3.resource('s3')
conn = boto3.client(
    's3',
    aws_access_key_id=config['AWS']['aws_access_key_id'],
    aws_secret_access_key=config['AWS']['aws_secret_access_key']
)
response = conn.list_buckets()

####### Function Started ############
def listBuckets():
    for bucket in response['Buckets']:
        print(bucket['Name'])


# Download Survey From AWS S3 Path, RAW Data
def downloadSurveyAnswer():
    objs = conn.list_objects(Bucket=config['AWS']['bucket'],
                             Prefix=config['AWS']['prefix'], Delimiter='/')
    if 'Contents' in objs:
        objs_contents = objs['Contents']
        print('✅ Start Downloading Process, Client: ' + client)
        for i in range(len(objs_contents)):
            filenameWithPath = objs_contents[i]['Key']
            FilenameArray = filenameWithPath.split('/')
            try:
                conn.download_file(config['AWS']['bucket'], filenameWithPath,
                                   downloadFolder+'/' + FilenameArray[-1])
                print('  ☑ File downloaded: ' + FilenameArray[-1])
            except:
                print('  ❌ Download Error!' + FilenameArray[-1])
        print('✅ Number of files downloaded: 【' + str(len(objs_contents)) + '】')
        print('✅ Start Extraction Process')

# Extract Survey Template
def surveyExtraction():
    shutil.rmtree(tempFolder)
    os.mkdir(tempFolder)
    print('✅ Start Extracting Survey From Template')
    global questionID_en, questions_en, questionID_zh, questions_zh, questionID_fr, questions_fr
    if len(os.listdir(inputFolder)) == 0:
        print("⚠️ Input Directory is empty")
        try:
            for inputFileName in glob.glob1('./SurveyTemplates', '*.*'):
                if inputFileName.endswith('.xls'):
                    shutil.copy2('./SurveyTemplates/' + inputFileName,
                                 inputFolder + '/' + inputFileName)
                    print("✅ Copied File : " + inputFileName)
        except:
            print("❌ Copy File Failed")
    else:
        print("✅ Input Directory is not empty")

    # if No files in Input folder, will automatically copy files into Input folder
    # then execute next function
    for inputFileName in glob.glob1(inputFolder, '*.*'):
        if inputFileName.endswith('.xls'):
            with open(inputFolder + '/' + inputFileName, 'r') as f:
                if (inputFileName == 'SurveyAnswers_en_template.xls'):
                    data = xlrd.open_workbook(inputFolder + '/' + inputFileName)
                    table = data.sheet_by_index(0)
                    questionID_en = table.col_values(0)
                    questions_en = table.col_values(1)
                    for idx, val in enumerate(questionID_en):
                        sheet_en.write(idx, 0, questionID_en[idx], style=QuestionStyle)
                    for idx, val in enumerate(questions_en):
                        sheet_en.write(idx, 1, questions_en[idx], style=QuestionStyle)
                elif (inputFileName == 'SurveyAnswers_fr_template.xls'):
                    data = xlrd.open_workbook(inputFolder + '/' + inputFileName)
                    table = data.sheet_by_index(0)
                    questionID_fr = table.col_values(0)
                    questions_fr = table.col_values(1)
                    for idx, val in enumerate(questionID_fr):
                        sheet_fr.write(idx, 0, questionID_fr[idx], style=QuestionStyle)
                    for idx, val in enumerate(questions_fr):
                        sheet_fr.write(idx, 1, questions_fr[idx], style=QuestionStyle)
                elif (inputFileName == 'SurveyAnswers_zh_template.xls'):
                    data = xlrd.open_workbook(inputFolder + '/' + inputFileName)
                    table = data.sheet_by_index(0)
                    questionID_zh = table.col_values(0)
                    questions_zh = table.col_values(1)
                    for idx, val in enumerate(questionID_zh):
                        sheet_zh.write(idx, 0, questionID_zh[idx], style=QuestionStyle)
                    for idx, val in enumerate(questions_zh):
                        sheet_zh.write(idx, 1, questions_zh[idx], style=QuestionStyle)
    print('✅ Finished Survey Template Manipulation')

    print('✅ Start Extracting JSON')
    for inputFileName in glob.glob1(downloadFolder, '*.*'):
        if inputFileName.endswith('.json'):
            with open(downloadFolder + '/' + inputFileName, 'r') as f:
                for lines in f:
                    match = re.search('(\{.*\})', lines)
                    if match:
                        surveyGroups = match.group()
                        for surveyGroup in match.groups():
                            f = open('./temp/RawData.json', 'a')
                            f.write(str(surveyGroup) + ",")
                            f.close()
                            
                            with open("./temp/RawData.json", "rt") as finStep1:
                                with open("./temp/out.json", "w") as foutStep1:
                                    for line in finStep1:
                                        foutStep1.write("[" + line + "]")

                            # replace "},]" to "}]"
                            with open("./temp/out.json", "rt") as finStep2:
                                with open("./temp/out1.json", "w") as foutStep2:
                                    for line in finStep2:
                                        foutStep2.write(line.replace('},]', '}]'))
    
    print('✅ Finished JSON Formatting')
    # finally format done! this is a valid json file now
    data_str = open('./temp/out1.json').read()
    # convert to json_data
    json_data = json.loads(data_str)
    print('✅ Start Individual Survey Extraction')

    try:
        i = 0
        while i < len(json_data):
            # print(len(json_data))
            if json_data[i]['lang'] == 'en':
                questionID = questionID_en
                sheetName = sheet_en
            elif json_data[i]['lang'] == 'ch':
                questionID = questionID_zh
                sheetName = sheet_zh
            elif json_data[i]['lang'] == 'fr':
                questionID = questionID_fr
                sheetName = sheet_fr
            Write_toRawExcel(json_data, i, json_data[i]['lang'], questionID, sheetName)
            i += 1
    except:
        print('❌ NO Files in input Folder, Please Copy Template From SurveyTemplates Folder')

    print('✅ Valid JSON Data: ' + '【' + str(len(json_data)+1) + '】')
    book_en.save(SurveyAnswer_en_xls)
    book_zh.save(SurveyAnswer_zh_xls)
    book_fr.save(SurveyAnswer_fr_xls)

    # remove and create tempFolder again
    shutil.rmtree(tempFolder)
    os.mkdir(tempFolder)
    print('✅ Finished All Survey Extraction')

def formatDocuments():
    try:
        print('✅ Start Formatting Documents')
        df_en = pd.read_excel(SurveyAnswer_en_xls)
        df_zh = pd.read_excel(SurveyAnswer_zh_xls)
        df_fr = pd.read_excel(SurveyAnswer_fr_xls)
        df_en.columns = df_en.columns.str.replace('Unnamed:', 'Survey Answer ')
        df_zh.columns = df_zh.columns.str.replace('Unnamed:', 'Survey Answer ')
        df_fr.columns = df_fr.columns.str.replace('Unnamed:', 'Survey Answer ')
        filteredData_en = df_en.dropna(axis='columns', how='all')
        filteredData_zh = df_zh.dropna(axis='columns', how='all')
        filteredData_fr = df_fr.dropna(axis='columns', how='all')
        writer = ExcelWriter(outputFolder + '/' + client + '_SurveyAnswer.xlsx')
        filteredData_en.to_excel(writer, 'en')
        filteredData_zh.to_excel(writer, 'zh')
        filteredData_fr.to_excel(writer, 'fr')
        writer.save()
        os.remove(SurveyAnswer_en_xls)
        os.remove(SurveyAnswer_zh_xls)
        os.remove(SurveyAnswer_fr_xls)
        print('✅ Finished Formatting Documents')
        print('✅ Files Have Been Combined into [' + client + '_SurveyAnswer.xlsx]')
    except:
        print('❌ Formatting Documents FAILED')


def SpreadSheetRowCount(workbook):
    book = xlrd.open_workbook(workbook)
    sheet = book.sheet_by_index(0)
    count = 0
    for row in range(sheet.nrows):
        count += 1
    return count

def Write_toRawExcel(json_data, i, languageKey, questionID, sheetName):
    j = 0
    if json_data[i]['lang'] == languageKey:
        while j < len(json_data[i]['surveyAnswers']):
            if json_data[i]['surveyAnswers'][j]['type'] == 'date' or json_data[i]['surveyAnswers'][j]['type'] == 'text' or json_data[i]['surveyAnswers'][j]['type'] == 'textarea':
                k = 0
                while k < len(questionID):
                    if questionID[k] == json_data[i]['surveyAnswers'][j]['id']:
                        sheetName.write(k, i+2, json_data[i]['surveyAnswers'][j]['value'], style=QuestionStyle)
                    k += 1
            elif json_data[i]['surveyAnswers'][j]['type'] == 'checkbox':
                k = 0
                while k < len(questionID):
                    for idx, val in enumerate(json_data[i]['surveyAnswers'][j]['value']):
                        if questionID[k] == val['id']:
                            sheetName.write(k, i+2, val['value'], style=QuestionStyle)
                    k += 1
            elif json_data[i]['surveyAnswers'][j]['type'] == 'radio':
                k = 0
                while k < len(questionID):
                    if questionID[k] == json_data[i]['surveyAnswers'][j]['value']['id']:
                        sheetName.write(k, i+2, json_data[i]['surveyAnswers'][j]['value']['value'], style=QuestionStyle)
                    k += 1
            j += 1

if __name__ == '__main__':
    downloadSurveyAnswer()
    surveyExtraction()
    formatDocuments()
