import os
import shutil
import xlrd
import xlwt
import xlutils.copy
from appJar import gui
import datetime
import configparser
import copy
from dealLine import *
from dealRepeat import *
from mysettings import *

def saveRecipe(myPara):
    file = myPara.sourceFile
    file_dir = os.path.dirname(file)
    file_name = os.path.basename(file)
    nowTime=datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    dst=os.path.join(file_dir,os.path.splitext(file_name)[0]+" "+nowTime+myPara.engineerName+os.path.splitext(file_name)[1])
    shutil.copy(file,dst)
    
    return(dst)

def processRecipe(myPara):
    if not myPara.sourceFile:
        app.warningBox('警告',"请提供菜单源文件!")
        return
    recipe = saveRecipe(myPara)

    #读取文件，生长速率调整并返回数据表，HCL位置
    [dataList,nrows,position] = adjustGrowthRate(recipe,myPara)
    #将修改的数据写入文件
    writeModefiedGrowthRateToFile(recipe,dataList,nrows,position) 

    while not stopLoop(dataList):
        if not processData(app,dataList,myPara):
            app.errorBox("错误","菜单错误，终止执行")
            return
        
    exportRecipe(dataList,myPara)

    app.infoBox("恭喜","完成！！！！！！！！！！")
    
def exportRecipe(dataList,myPara):    
    finalList=[]
    lineList=[]
    #整理表格
    #第一行
    lineList.append("step")
    for i in dataList[0][4:]:
        lineList.append(i)
    temp = copy.deepcopy(lineList)
    finalList.append(temp)
    lineList.clear()
    #余下行
    j=1
    for i in dataList[1:]:
        lineList.append(j)
        for k in range(4,len(i)):
            lineList.append(i[k])
        temp = copy.deepcopy(lineList)
        finalList.append(temp)
        lineList.clear()
        j=j+1

    tTime = totalTime(finalList)
    message = "预计菜单执行时间总计(h):"+str(round(tTime/3600,2))
    app.infoBox("执行成功",message)

    #写入Excel
    writeWorkBook=xlwt.Workbook(encoding="utf-8")
    writeTable=writeWorkBook.add_sheet("recipe")
    for i in range(len(finalList)):
        for j in range(len(finalList[0])):
            writeTable.write(i,j,finalList[i][j])

    nowTime=datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    file = os.path.join(myPara.exportFolder,nowTime+myPara.engineerName+".xls")        
    writeWorkBook.save(file)

#读取文件，生长速率调整并返回数据表
def adjustGrowthRate(recipe,myPara):
    readWorkBook=xlrd.open_workbook(recipe)
    readTable=readWorkBook.sheet_by_index(0)

    nrows=readTable.nrows
    ncols=readTable.ncols
    dataList=[]
    for i in range(nrows):
        dataList.append(readTable.row_values(i))
    
    position = dataList[0].index("HCL")
    growthRateFactor = myPara.growthRateFactor
    for i in dataList[1:]:
        i[position]=round(i[position]*growthRateFactor)

    return [dataList,nrows,position]

def writeModefiedGrowthRateToFile(file,dataList,nrows,position):
    rb = xlrd.open_workbook(file)
    wb = xlutils.copy.copy(rb)
    ws = wb.get_sheet(0)
    for i in range(1,nrows):
        ws.write(i,position,dataList[i][position])
    wb.save(file)

def totalTime(list):
    position = list[0].index("step time")
    time = 0
    for i in range(1,len(list)):
        time = time+list[i][position]
    return time    

def processData(app,dataList,myPara):
    #处理线性变化
    sucess01 = delLine(app,dataList,myPara)
    #处理周期变化
    sucess02 = delRepeat(app,dataList)

    finalSucess = sucess01 and sucess02
    return finalSucess  
        
#判断终止循环
def stopLoop(dataList):
    flag=True
    RNumber=0
    for i in dataList:
        if i[0]=="L" or i[0]=="R":
            flag = False
    return flag         

def displayDataList(dataList):
    for i in dataList:
        print(i)

def press_select(button):
    if button=="button1":
        temp = app.directoryBox("选择文件夹")
        if temp:
            app.setEntry("菜单导出路径", temp)
    if button=="button2":
        temp = app.directoryBox("选择文件夹")
        if temp:
            app.setEntry("源菜单路径", temp)
    if button=="button3":
        temp = app.openBox("选择文件",dirName=app.getEntry("源菜单路径"),fileTypes=[("excelFiles",".xls"),("excelFiles",".xlsx")])
        if temp:
            app.setEntry("源菜单", temp)

def press_action(button):
    if button=="开始":
        myPara = constructParameter(app)
        processRecipe(myPara)
    if button=="清空":
        app.clearEntry("菜单导出路径")
        app.clearEntry("源菜单路径")
    if button=="保存":
        stepTime=str(app.getEntry("默认步间距"))
        growthRateFactor=str(app.getEntry("HCL流量因子"))
        enginearName=app.getEntry("责任工程师")
        exfileFolder=app.getEntry("菜单导出路径")
        sFolder=app.getEntry("源菜单路径")
        file=app.getEntry("源菜单")
        saveSetting(stepTime,growthRateFactor,enginearName,exfileFolder,sFolder,file) 

app = gui("HVPE菜单生成器")

valueList=readSetting()

app.addLabel("lb1","默认步间距",0,0)
app.addNumericEntry("默认步间距",0,1)
app.setEntry("默认步间距",valueList[0])

app.addLabel("lb2","HCL流量因子",1,0)
app.addNumericEntry("HCL流量因子",1,1)
app.setEntry("HCL流量因子",valueList[1])

app.addLabelEntry("责任工程师",2,0)
app.setEntry("责任工程师",valueList[2])

app.addLabelEntry("菜单导出路径",3,0)
app.setEntry("菜单导出路径",valueList[3])
app.addNamedButton("选择","button1",press_select,3,1)

app.addLabelEntry("源菜单路径",4,0)
app.addNamedButton("选择","button2",press_select,4,1)
app.setEntry("源菜单路径",valueList[4])

app.addLabelEntry("源菜单",5,0)
app.setEntry("源菜单",valueList[5])
app.addNamedButton("选择","button3",press_select,5,1)

app.addButtons(["开始", "清空","保存"], press_action,6,0,2)
app.go()
