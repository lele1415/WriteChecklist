#!/usr/bin/env python
#coding=utf-8
# Date: 2017/08/31
# Author: chenxs

import sys
reload(sys)
sys.setdefaultencoding('utf8')

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
            sLine = sLine.replace(chr(32), "")
            sLine = sLine.replace(chr(9), "")
            sLine = sLine.replace("\n", "")
            if re.match("^" + sKey + "=", sLine):
                sLine = sLine.split("=")
                return sLine[1]
    elif sMode == modeItems:
        if sKey in sLine:
            sLine = sLine.replace(chr(9), chr(32))
            sLine = sLine.replace("\n", "")
            if re.match("^" + sKey, sLine):
                while "  " in sLine:
                    sLine = sLine.replace(chr(32) + chr(32), chr(32))
                sLine = sLine.split(chr(32))
                return sLine[1]
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
modem = ""
if checkFileExist(sProjectConfigPath_opt, False):
    modem = getValueInFile(sProjectConfigPath_opt, "CUSTOM_MODEM", modeSystemprop)
    findInPrjFlag = (modem == "")
else:
    findInPrjFlag = True

if findInPrjFlag:
    if checkFileExist(sProjectConfigPath_prj, False):
        modem = getValueInFile(sProjectConfigPath_prj, "CUSTOM_MODEM", modeSystemprop)

###############
# write excel
###############
import time
date = time.strftime('%Y/%m/%d',time.localtime(time.time()))
import commands
(status, author) = commands.getstatusoutput("git config --global user.name")

import xlsxwriter
workbook = xlsxwriter.Workbook('checklist_tmp.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 60)
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

formatBlue = setInfoValueFormat(True, 'blue')
formatOrange = setInfoValueFormat(False, 'E26B0A')
formatRed = setInfoValueFormat(True, 'red')
formatGreen = setInfoValueFormat(True, '76933C')
formatPink = setInfoValueFormat(True, 'pink')
formatGray = setInfoKeyFormat('808080')

worksheet.write('A3', date, formatDateAndAuthor)
worksheet.write('A4', author, formatDateAndAuthor)

worksheet.write('B3', "版本号: ", formatGray)
worksheet.write('B4', "项目-客户: ", formatGray)
worksheet.write('B5', "工程: ", formatGray)
worksheet.write('B6', "屏驱动: ", formatGray)
worksheet.write('B7', "TP驱动: ", formatGray)
worksheet.write('B8', "modem: ", formatGray)
worksheet.write('B9', "提交节点: ", formatGray)
worksheet.write('B10', "分支名: ", formatGray)
worksheet.write('B11', "代码路径: ", formatGray)
worksheet.write('B12', "禅道号: ", formatGray)

worksheet.write('C3', displayId, formatBlue)
worksheet.write('C4', roco_project + "-" + customName, formatRed)
worksheet.write('C5', target_product, formatRed)
worksheet.write('C6', lcm, formatPink)
worksheet.write('C7', touchpanel, formatPink)
worksheet.write('C8', modem, formatPink)
worksheet.write('C9', commitId, formatOrange)
worksheet.write('C10', branch, formatGreen)
worksheet.write('C11', pwd, formatGreen)
worksheet.write('C12', taskNum, formatGreen)
workbook.close()