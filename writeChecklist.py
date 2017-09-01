#!/usr/bin/env python
#coding=utf-8
# Version: v1.0.2
# Date: 2017/09/01
# Author: chenxs

import sys
reload(sys)
sys.setdefaultencoding('utf8')

def red(sText):
    return '\033[0;31m ' + sText + ' \033[0m'
def green(sText):
    return '\033[0;32m ' + sText + ' \033[0m'
def yellow(sText):
    return '\033[0;33m ' + sText + ' \033[0m'
def greenAndYellow(sText1, sText2):
    return green(sText1) + yellow(sText2)

#################
# get infos
#################
import commands
import sys
import os

# get custom name
customName = raw_input('客户名：')
taskNum = raw_input('禅道号：')

# get branch
(status, branchs) = commands.getstatusoutput("git branch")
branchList = branchs.split("\n")
for br in branchList:
    if "*" in br:
        branch = br.replace("* ", "")
        break

# get pwd
(status, pwd) = commands.getstatusoutput("pwd")

# get last commit id
(status, log) = commands.getstatusoutput("git log -1")
commitId = log.split("\n")
commitId = commitId[0].replace("commit ", "")

# get project info
(status, roco_project) = commands.getstatusoutput("echo $ROCO_PROJECT")
(status, target_product) = commands.getstatusoutput("echo $TARGET_PRODUCT")
target_product = target_product.replace("full_", "")
if (roco_project == "") or (target_product == ""):
    print("请先lunch")
    sys.exit(0)
#print("ROCO_PROJECT = " + roco_project + "\nTARGET_PRODUCT = " + target_product)

# get display_id, lcm, tp
def checkFileExist(filePath, exitFlag):
    if not os.path.exists(filePath):
        if exitFlag:
            print("File is not exist:\n" + filePath)
            sys.exit(0)
        else:
            return False
    return True

sItemsPath = "device/joya_sz/" + target_product + "/roco/" + roco_project + "/items.ini"
sSystempropPath = "device/joya_sz/" + target_product + "/roco/" + roco_project + "/system.prop"
sProjectConfigPath_opt = "device/joya_sz/" + target_product + "/roco/" + roco_project + "/ProjectConfig.mk"
sProjectConfigPath_prj = "device/joya_sz/" + target_product + "/ProjectConfig.mk"

checkFileExist(sItemsPath, True)
checkFileExist(sSystempropPath, True)

import re

modeSystemprop = 0
modeItems = 1

def getValueInLine(sLine, sKey, sMode):
    if sMode == modeSystemprop:
        if sKey in sLine:
            sLine = sLine.strip()
            if sLine.startswith(sKey):
                sLine = sLine.split("=")
                return sLine[1].strip()
    elif sMode == modeItems:
        if sKey in sLine:
            sLine = sLine.strip()
            sLine = sLine.replace(chr(9), " ")
            if sLine.startswith(sKey):
                while "  " in sLine:
                    sLine = sLine.replace("  ", " ")
                sLine = sLine.split(" ")
                return sLine[1].strip()
    return ""

def getValueInFile(sFilePath, sKey, sMode):
    file = open(sFilePath, 'r')
    lines = file.readlines()
    for line in lines:
        sValue = getValueInLine(line, sKey, sMode)
        if sValue != "":
            return sValue
    return ""
    file.closed


displayId = getValueInFile(sSystempropPath, "ro.build.display.id", modeSystemprop)
lcm = getValueInFile(sItemsPath, "LCM", modeItems)
touchpanel = getValueInFile(sItemsPath, "touchpanel.gsl.modle", modeItems)

if ("8127" in pwd) or ("8163" in pwd) or ("8167" in pwd):
    isWifiPlatform = True
else:
    isWifiPlatform = False

if not isWifiPlatform:
    modem = ""
    if checkFileExist(sProjectConfigPath_opt, False):
        modem = getValueInFile(sProjectConfigPath_opt, "CUSTOM_MODEM", modeSystemprop)
        findInPrjFlag = (modem == "")
    else:
        findInPrjFlag = True
    
    if findInPrjFlag:
        if checkFileExist(sProjectConfigPath_prj, False):
            modem = getValueInFile(sProjectConfigPath_prj, "CUSTOM_MODEM", modeSystemprop)
    
    if (modem != "") and (" " in modem):
        multiModemFlag = True
        modem = modem.split(" ")
    else:
        multiModemFlag = False

###############
# write excel
###############
import time
date = time.strftime('%Y/%m/%d',time.localtime(time.time()))
import commands
(status, author) = commands.getstatusoutput("git config --global user.name")

checklistDirPath = '../Checklist'
checklistFilePath = '../Checklist/checklist_tmp.xlsx'
if not os.path.exists(checklistDirPath):
    os.makedirs(checklistDirPath)
if os.path.exists(checklistFilePath):
    os.remove(checklistFilePath)

import xlsxwriter
workbook = xlsxwriter.Workbook(checklistFilePath)
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 70)
worksheet.set_column('D:D', 60)

formatUpdateInfo = workbook.add_format()
formatUpdateInfo.set_font_color('31869B')
formatUpdateInfo.set_align('top')
formatUpdateInfo.set_bold()

worksheet.merge_range('D3:D12', '修改点见禅道' + taskNum, formatUpdateInfo)

formatDateAndAuthor = workbook.add_format()
formatDateAndAuthor.set_bottom(1)
formatDateAndAuthor.set_top(1)
formatDateAndAuthor.set_left(1)
formatDateAndAuthor.set_right(1)
formatDateAndAuthor.set_bold()
formatDateAndAuthor.set_align('center')
formatDateAndAuthor.set_font_color('red')
formatDateAndAuthor.set_font_size(12)
formatDateAndAuthor.set_bg_color('yellow')

def setInfoKeyFormat(color):
    formatObj = workbook.add_format()
    formatObj.set_bottom(1)
    formatObj.set_top(1)
    formatObj.set_left(1)
    formatObj.set_right(1)
    formatObj.set_bold()
    formatObj.set_align('right')
    formatObj.set_font_color(color)
    formatObj.set_font_size(12)
    return formatObj

def setInfoValueFormat(setBold, color):
    formatObj = workbook.add_format()
    formatObj.set_bottom(1)
    formatObj.set_top(1)
    formatObj.set_left(1)
    formatObj.set_right(1)
    if setBold:
        formatObj.set_bold()
    formatObj.set_align('center')
    formatObj.set_font_color(color)
    formatObj.set_font_size(12)
    return formatObj

def writeCellValue(sValue, oFormat):
    global rowNum
    worksheet.write(columnNum + rowNum, sValue, oFormat)
    rowNum = str(int(rowNum) + 1)

formatBlue = setInfoValueFormat(True, 'blue')
formatOrange = setInfoValueFormat(False, 'E26B0A')
formatRed = setInfoValueFormat(True, 'red')
formatGreen = setInfoValueFormat(True, '76933C')
formatPink = setInfoValueFormat(True, 'pink')
formatGray = setInfoKeyFormat('808080')

columnNum = 'A'
rowNum = '3'
writeCellValue(date, formatDateAndAuthor)
writeCellValue(author, formatDateAndAuthor)

columnNum = 'B'
rowNum = '3'
writeCellValue("版本号: ", formatGray)
writeCellValue("项目-客户: ", formatGray)
writeCellValue("工程: ", formatGray)
writeCellValue("屏驱动: ", formatGray)
writeCellValue("TP驱动: ", formatGray)
if not isWifiPlatform:
    if not multiModemFlag:
        writeCellValue("modem: ", formatGray)
    else:
        modemCount = 0
        for md in modem:
            modemCount = modemCount + 1
            writeCellValue("modem" + str(modemCount) + ": ", formatGray)
writeCellValue("提交节点: ", formatGray)
writeCellValue("分支名: ", formatGray)
writeCellValue("代码路径: ", formatGray)
writeCellValue("禅道号: ", formatGray)

columnNum = 'C'
rowNum = '3'
writeCellValue(displayId, formatBlue)
writeCellValue(roco_project + "-" + customName, formatRed)
writeCellValue(target_product, formatRed)
writeCellValue(lcm, formatPink)
writeCellValue(touchpanel, formatPink)
if not isWifiPlatform:
    if not multiModemFlag:
        writeCellValue(modem, formatPink)
    else:
        for md in modem:
            writeCellValue(md, formatPink)
writeCellValue(commitId, formatOrange)
writeCellValue(branch, formatGreen)
writeCellValue(pwd, formatGreen)
writeCellValue(taskNum, formatGreen)

workbook.close()

print(red("########## start ##########"))
print(greenAndYellow("版本号: ", displayId))
print(greenAndYellow("项目-客户:", roco_project + "-" + customName))
print(greenAndYellow("工程: ", target_product))
print(greenAndYellow("屏驱动: ", lcm))
print(greenAndYellow("TP驱动: ", touchpanel))
if not isWifiPlatform:
    if not multiModemFlag:
        print(greenAndYellow("modem: ", modem))
    else:
        modemCount = 0
        for md in modem:
            modemCount = modemCount + 1
            print(greenAndYellow("modem" + str(modemCount) + ": ", md))
print(greenAndYellow("提交节点: ", commitId))
print(greenAndYellow("分支名: ", branch))
print(greenAndYellow("代码路径: ", pwd))
print(greenAndYellow("禅道号: ", taskNum))
print(red("########## end ##########"))