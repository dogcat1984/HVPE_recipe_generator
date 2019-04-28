import copy as cp

#处理线性变化
def delLine(app,dataList,myPara):
    sucess = True
    minStepTime = myPara.lineStepTime
    #元素个数
    n = len(dataList)
    #由第三行开始
    for i in range(2,n):
        line = dataList[i]
        if line[0]=="L" and dataList[i-1]:#如果当前行标记为L且上一行不为空
            
            if line[2]:#如果指定线性行数
                try:
                    stepTime = int(line[5]/line[2])#重新计算步间距
                except:                    
                    message = "数据类型错误，位置（行）: "+str(i+1)
                    app.errorBox("菜单错误",message)
                    sucess = False
                    return sucess
            else:
                stepTime = minStepTime#如果没指定线性行数，将每步时间指定为默认步间距

            if stepTime>=minStepTime:#步间距必须大于默认步间距，才执行
                firstRows=int(line[5]/stepTime)
                factor=stepTime/line[5]
                remainder=line[5]%stepTime#时间对每步时间求余数
                #插入线性行，参数为步间距，上一行，当前行，插入的行数，数据源，插入行起始位置，因子
                insertLineLine(stepTime,dataList[i-1],line,firstRows,dataList,i,factor)

                #余数不为0
                if remainder!=0:
                    insertModLineLine(remainder,firstRows,dataList,i)
                #删除当前行
                dataList.pop(i)

            else:
                message = "时间间距过小，请检查菜单设置!错误发生位置（行）："+str(i+1)
                app.errorBox("菜单错误",message)
                sucess = False
    return sucess
    
#参数为步间距，上一行，当前行，插入的行数，数据源，插入行起始位置，因子                
def insertLineLine(stepTime,sline,eline,newRows,dataList,position,factor):
    for i in range(newRows):
        newLine=[]
        newLine.append("")
        newLine.append("")
        newLine.append("")
        newLine.append("")
        for j in range(4,len(sline)):     
            value=round(sline[j]+(eline[j]-sline[j])*factor*(i+1))
            newLine.append(value)
        newLine[5]=stepTime
        dataList.insert(position+1+i,newLine)

#参数为插入的行数，数据源，插入行起始位置
def insertModLineLine(remainder,newRows,dataList,position):
    newLine=[]
    newLine.append("")
    newLine.append("")
    newLine.append("")
    newLine.append("")
    newLine.append(cp.deepcopy(dataList[position][4]))
    newLine.append(remainder)

    for j in range(6,len(dataList[position])):
        value=dataList[position][j]
        newLine.append(cp.deepcopy(value))            
    dataList.insert(position+newRows+1,newLine)
