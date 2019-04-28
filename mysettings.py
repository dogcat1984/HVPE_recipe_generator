import configparser

class parameter:
    def __init__(self,lst,gRF,eN,exFo,sFo,sF):
        self.lineStepTime=lst
        self.growthRateFactor=gRF
        self.engineerName=eN        
        self.exportFolder=exFo
        self.sourceFolder=sFo
        self.sourceFile=sF

def saveSetting(stepTime,growthRateFactor,enginearName,exfileFolder,sFolder,file):
    conf = configparser.ConfigParser()
    conf.add_section("config")
    conf.set("config","默认步间距",stepTime)
    conf.set("config","HCL流量因子",growthRateFactor)
    conf.set("config","责任工程师",enginearName)
    conf.set("config","菜单导出路径",exfileFolder)
    conf.set("config","源菜单路径",sFolder)
    conf.set("config","源菜单",file)

    with open("conf.ini","w") as fw:
        conf.write(fw)

def readSetting():
    conf = configparser.ConfigParser()
    valueList=[]
    conf.read("conf.ini")
    valueList.append(conf.get("config","默认步间距"))
    valueList.append(conf.get("config","HCL流量因子"))
    valueList.append(conf.get("config","责任工程师"))
    valueList.append(conf.get("config","菜单导出路径"))
    valueList.append(conf.get("config","源菜单路径"))
    valueList.append(conf.get("config","源菜单"))
    return valueList

def constructParameter(app):
    stepTime=app.getEntry("默认步间距")
    growthRateFactor=app.getEntry("HCL流量因子")
    enginearName=app.getEntry("责任工程师")
    exfileFolder=app.getEntry("菜单导出路径")
    sFolder=app.getEntry("源菜单路径")
    file=app.getEntry("源菜单")
    myPara = parameter(stepTime,growthRateFactor,enginearName,exfileFolder,sFolder,file)
    return myPara
