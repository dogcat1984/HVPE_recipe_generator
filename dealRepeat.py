#处理周期变化
def delRepeat(app,dataList):
    sucess = True
    n = len(dataList)
    #由第二行
    labelList=[]
    for i in range(1,n):
        line = dataList[i]
        if line[0]=="R":
            labelList.append([line[1],line[2],i])#循环名、循环次数、行数
    if labelList:
        l = len(labelList)
        if countList(app,labelList):
            sucess = False
        else:
            for i in range(l-1):
                if labelList[i][0]==labelList[i+1][0] and labelList[i][1]==labelList[i+1][1]:
                    repeatTimes=labelList[i][1]
                    try:
                        int(repeatTimes)
                    except:
                        message = "循环次数设置错误！位置："+str(labelList[i][2]+1)
                        app.errorBox("菜单错误",message)
                        sucess = False
                    else:
                        startRow=labelList[i][2]#开始循环位置
                        endRow=labelList[i+1][2]#结束循环位置
                        block=decideBlock(startRow,endRow,dataList)#确定循环块
                        insertRepeatLine(block,dataList,endRow,repeatTimes)#插入循环块
                        clearOriginBlock(startRow,endRow,dataList)
                if labelList[i][0]==labelList[i+1][0] and labelList[i][1]!=labelList[i+1][1]:
                    message = "循环次数设置错误，循环名为 "+str(labelList[i][0])
                    app.errorBox("菜单错误",message)
                    sucess = False
    return sucess

#参数为周期块，数据源，插入行起始位置，周期数            
def insertRepeatLine(block,dataList,endRow,times):
    k=endRow
    for i in range(int(times)):
        for j in range(len(block)):
            k=k+1
            dataList.insert(k,block[j])
    
#确定循环块            
def decideBlock(sRow,eRow,dataList):
    block=[]
    headLine=[]
    tailLine=[]

    #循环块起始行
    headLine.append("")
    headLine.append("")
    headLine.append("")
    headLine.append("")
    for i in range(4,len(dataList[sRow])):
        headLine.append(dataList[sRow][i])
    block.append(headLine)    

    #循环块中间行
    for i in range(sRow+1,eRow):
        block.append(dataList[i])

    #循环块尾行               
    tailLine.append("")
    tailLine.append("")
    tailLine.append("")
    tailLine.append("")
    for i in range(4,len(dataList[eRow])):
        tailLine.append(dataList[eRow][i])
    block.append(tailLine)

    return block

#清除原始循环块
def clearOriginBlock(sRow,eRow,dataList):
    for i in range(eRow-sRow+1):
        dataList.pop(sRow)

#统计循环名是否缺失
def countList(app,labelList):
    #获取labelList中元素里的循环名
    tempList=[]
    for i in labelList:
        tempList.append(i[0])
    #tempList去重
    news_ids = []
    for id in tempList:
        if id not in news_ids:
            news_ids.append(id)
    #统计        
    for i in news_ids:
        result = tempList.count(i)
        if result%2==1:
            message = "周期变化标记缺失！！！！！！！！！！\n"+"缺失周期为"+str(i[0])
            app.errorBox("菜单错误",message)
            return True
