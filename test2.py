from asyncio.windows_events import NULL
from openpyxl import load_workbook
import  os
from itertools import zip_longest

rowNum = 0
colNum = 0
backUpDirectory = ''
excel = NULL

def setFilePath (filePath) :
    if (filePath != '') :
        return filePath

def getSheetNames(excel) :
    # book = xlrd.open_workbook(backUpDirectory)
    print("The number of worksheets is {0}".format(len(excel.sheetnames)))
    print("Worksheet name(s): {0}".format(excel.sheetnames))
    return excel.sheetnames

def getSheetDataList(excel,sheetName) :
    sqlList = []
    table = excel.get_sheet_by_name(sheetName)
    # Data=table.cell(row=rowNum,column=colNum).value
    for i in range(3,table.max_row+1): #从2开始，第一行为title
        #遍历Excel表格列的内容
        sqlData = (table.cell(row=i,column=1)).value #获取数据
        if ((sqlData != '') and not(sqlData is None)) :
            sqlList.append(sqlData)
            print("读取到的SQL文：%s",sqlData)
    return sqlList

# 指定规则：表项目取得
# 返回值  1:读取Excel Sheet的一个位置的表明 String
#        2：读取Excel Sheet的第一行的各项目名 List
def getSheetHeadDataList(excel,sheetName) :
    sqlTableName = ''
    sqlHead = []
    table = excel.get_sheet_by_name(sheetName)
    sqlTableName = (table.cell(row=1,column=1)).value #获取数据
    # Data=table.cell(row=rowNum,column=colNum).value
    for i in range(4,table.max_column+1): #从3开始，第一行为title
        #遍历Excel表格列的内容
        sqlData = (table.cell(row=2,column=i)).value #获取数据
        if ((sqlData != '') and not(sqlData is None)) :
            sqlHead.append(sqlData)
            # print("读取到的SQL项目名：%s",sqlHead)
    return sqlTableName,sqlHead

# 指定规则：各项目确认内容取得
def getSheetBodyDataList(excel,sheetName) :
    sqlBodyList = []
    table = excel.get_sheet_by_name(sheetName)
    # nullCheckList = [False,False]
    for x in range(3,table.max_row+1): #从2开始，第一行为title
    
        sqlBodyAllList = []
        # Data=table.cell(row=rowNum,column=colNum).value
        nullCheck = False
        for i in range(4,table.max_column+1): #从3开始，第一行为title
            #遍历Excel表格列的内容
            sqlData = (table.cell(row=x,column=i)).value #获取数据
            if (not(sqlData is None)) :
                nullCheck = True
                sqlBodyAllList.append(sqlData)
                # print("读取到的SQL文各项目值：%s",sqlBodyAllList)
            else:
                # Excel空的时候，值为None
                sqlBodyAllList.append('')
                # nullCheckList.append()
                # print("读取到的SQL文各项目值错误")
        # 一行的数据不全为空的数据
        if (nullCheck) :
            sqlBodyList.append(sqlBodyAllList)
    return sqlBodyList
    
def is_kazu(num):
    if type(num) is not str:
        raise ValueError('parameter must be a string.')
    if len(num) > 2 and num.count('.',1,-1) == 1:
        num=num.replace('.','',1)
    if num.isnumeric():
        return True
    else:
        return False
    
def openExecl(paraUpDirectory):
    return load_workbook(paraUpDirectory, data_only=True)

def getInputExcel(excel, sheetName):
    sqlBodyList = []
    table = excel.get_sheet_by_name(sheetName)
    # nullCheckList = [False,False]
    for x in range(1,table.max_row+1): #从0开始，第一行为title
    
        sqlBodyAllList = {}
        # Data=table.cell(row=rowNum,column=colNum).value
        nullCheck = False
        for i in range(1,table.max_column+1): #从0开始，第一行为title
            #遍历Excel表格列的内容
            ruleData = (table.cell(row=x,column=i)).value #获取数据
            if (not(ruleData is None)) :
                nullCheck = True
                sqlBodyAllList.append(ruleData)
                # print("读取到的SQL文各项目值：%s",sqlBodyAllList)
            else:
                # Excel空的时候，值为None
                sqlBodyAllList.append('')
                # nullCheckList.append()
                # print("读取到的SQL文各项目值错误")
        # 一行的数据不全为空的数据
        if (nullCheck) :
            sqlBodyList.append(sqlBodyAllList)
    return sqlBodyList
    print('')
    
def readExcelmain(paraUpDirectory):
    if (paraUpDirectory == '') :
        #得到当前脚本的执行目录
        currentDirectory  =  os.getcwd()
        #查看是否已经存在备份目录，如果有则删除，没有则新建目录
        paraUpDirectory  =  "%s\\%s"  %( currentDirectory, "rule.xlsx")
    excel=load_workbook(paraUpDirectory, data_only=True)
    
    sheetNameList = getSheetNames(excel)
    for sheetName in sheetNameList:
        if (sheetName == 'InputExcel') :
            print(sheetName)
            sqlList = getInputExcel(excel, str(sheetName))

def main(paraUpDirectory,excel):
    if (paraUpDirectory == '') :
        #得到当前脚本的执行目录
        currentDirectory  =  os.getcwd()
        #查看是否已经存在备份目录，如果有则删除，没有则新建目录
        paraUpDirectory  =  "%s\\%s"  %( currentDirectory, "SPUR Phase3_メンテンマスタ初期値設定リスト.xlsx")
    excel=load_workbook(paraUpDirectory, data_only=True)
    
    sheetNameList = getSheetNames(excel)
    for sheetName in sheetNameList:
        if (is_kazu(sheetName)) :
            print(sheetName)
            sqlList = getSheetDataList(excel, str(sheetName))


def getExcelSheetDatalist(excel,sheetName):
    ruleBodyList = {}
    table = excel.get_sheet_by_name(sheetName)
    # nulCheckList =[False,False]
    for x in range(1,table.max_row+1): #从2开始，第一行为title
        # 带遍烟历Excel表格列的内容
        rowData =(table.cell(row=x,column=1)).value #获取数据
        rulecolBodyAllList= []
        nullcheck = False
        for i in range(2,table.max_column+1): #从3开始，第一行为title
            # 看历Excel表格列的内容
            colData = (table.cell(row=x,column=i)).value #获取数据
            if(not(colData is None)):
                nullcheck = True
                rulecolBodyAllList.append(colData)
        # 一行的数据不全为空的数据
        if(nullcheck):
            ruleBodyList[rowData]=rulecolBodyAllList 
    return ruleBodyList

def setExcelData(ruleList, outList):
    # 取得Input文件名
    inputFilePath = ruleList['filename']
    # 取得Output文件名
    outputFilePath = outList['filename']
    # 取得数据规则Head
    datarule = ruleList['datarule']
    # Input信息取得
    workSheetList = []
    itemNameList = []
    itemList = []
    for rowstr in ruleList:
        # 取得Input Sheet名
        inputData = ruleList[rowstr]
        if 'workSheet' in rowstr:
            workSheetList.append(inputData)
            print('')
        if 'item' in rowstr:
            itemNameList.append(rowstr)
            itemList.append(inputData)

    # Output信息取得
    outWorkSheetList = []
    outItemNameList = []    
    outItemList = []    
    for outstr in outList:
        # 取得Input Sheet名
        outData = ruleList[outstr]
        if 'workSheet' in outstr:
            outWorkSheetList.append(outData)
            print('')
        if 'item' in outstr:
            outItemNameList.append(outstr)
            outItemList.append(outData)

    # 根据项目ID取得对应项目值 设定到对应Output项目值

def setItemData(datarule,workSheetList,itemNameList,itemList,outWorkSheetList,outItemNameList,outItemList):
    # 根据每个项目的取得信息，取得项目设定值
    # d1=zip_longest(l1,l2)
    item1index = {}
    item2index = {}
    item3index = {}
    item4index = {}
    for i,rulestr in enumerate(datarule):
        if ('rownum' == rulestr) :
            item1index

    print('')

# 获取项目信息，供后续处理调用
def getItemValue(datarule,itemNameList,itemList):
    itmeValue = {}
    for itemnmae,item in zip(itemNameList,itemList):
        itmeValue[itemnmae] = zip_longest(datarule,item)
        if (isinstance(itmeValue[itemnmae],list)):
            itemonelist = itmeValue[itemnmae]
            for itemone in itemonelist:
                if (itemonelist[itemone] == None):
                    del itemonelist[itemone]

    return itmeValue

# 'datarule': ['rownum', 'colnum', 'iftype', 'setvalue', 'condition', 'rownum', 'colnum', 'iftype', 'rownum', 'colnum']
# 'item3': [20, 3], 
# 'item4': [21, 3, 'and', ',', 'none', 23, 3, 'and', 24, 3], 
# 'item5': [22, 3, ' County'], 
# 'item6': [29, 3],

# 指定内容取得
def getData(table,rowNum,colNum) :
    strData = (table.cell(row=rowNum,column=colNum)).value #获取数据
    return strData


def main(paralpDirectory,excel):
    if(paralpDirectory == ''):
        # 得到当前脚本的执行目录
        cunnentDirectory = os.getcwd()
        # 希查看是否已经存在备份目录，如果有则删除，没有则新建目录
        paralpDirectory = "%s\\%s" %( cunnentDirectory,"doc\\input\\rule.xlsx")
    excel=load_workbook(paralpDirectory, data_only=True)

    sheetNameList = getSheetNames(excel)
    for sheetName in sheetNameList:
        if("excel"in sheetName):
            print(sheetName)
            rulelist=[]
            if('input'in sheetName):
                ruleList = getExcelSheetDatalist(excel,str(sheetName))
                print('ruleList')
                print(ruleList)
            outList = []
            if('output' in sheetName):
                outList = getExcelSheetDatalist(excel, str(sheetName))
                print('outList')
                print(outList)



if  __name__  ==  '__main__' :
    main('',excel)
    # getSheetDataList("1")

    print ("\n读取到完成...\n回车键退出")
