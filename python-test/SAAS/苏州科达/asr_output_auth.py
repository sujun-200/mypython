#!/usr/bin/python3
#-*- coding:utf-8 -*-
import requests
import sys
import json
import time
import os
import getopt
import codecs
import uuid
import queue
import threading
import wx
import hashlib
from openpyxl import Workbook
import traceback

con_excelName = 'asr_output.xlsx'
con_excelSheetName = '1'

con_productId = '914012188'
con_pubKey = '9035d19fea114efa9992cae90e5c55c3'
con_secKey = '16D7EDA2B84DC1E71CDC71B542D60720'
con_fenli = 'false'
con_hostName = 'https://api.talkinggenie.com'
con_upApi = '/smart/sinspection/api/v1/fileUpload'
con_getApi = '/smart/sinspection/api/v2/getTransResult'
con_niUrl = 'http://ezmt.duiopen.com/ezapi/inverse'


con_productId = '914013449'
con_pubKey = '15edc97f863747ea9fd1362b81909ab2'
con_secKey = '80E8A069936E904544279F049C4FFDCF'

con_fenli = 'true'
#con_hostName = 'http://10.12.7.233:19093'
#con_upApi = '/sinspection/api/v1/fileUpload'
#con_getApi = '/sinspection/api/v1/getTransResult'
#con_niUrl = 'http://10.12.7.233:40080/inverse'


TASK_STATUS_PRE = 0
TASK_STATUS_RUN = 1
TASK_STATUS_FIN = 2
TASK_STATUS_FAIL = 3

RETRY_TIMES = 3
g_filePath = ''
g_isNumberInverting = False
g_niUrl = 'http://ezmt.duiopen.com/ezapi/inverse'
g_niUrl = ''
g_finNum = 0

def numChg(oriText):
    if oriText == '\n':
        return oriText
    if oriText[-1] == '\n':
        oriText = oriText[:-1]

    dd = u'["'+oriText+u'"]'
    response = requests.request("POST", g_niUrl, 
                                headers={'Content-type':'application/json'}, data=dd.encode(encoding='utf8'))
    res = response.text.encode('utf-8') 
    j = json.loads(res)
    if j['errno']!=200 and j['errno']!=0:
        return oriText+'\n'
    aa = j['data']['result'][0]
    if aa[-1]!='\n':
        aa += '\n'
    return aa

def log(strLogContent):
    try:
        logFile = open("./log.log","ab+")
        preFix = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime()) + ": #### : "
        if isinstance(strLogContent,str):
            strLogContent = strLogContent.encode("utf-8")
        preFix = preFix.encode() + strLogContent + b"\n"
        logFile.write(preFix)
        logFile.close()
    except Exception as e:
        print (e)

class CAuthentication:
    def __init__(self,pid,publicKey,privateKey):
        self.pid = pid
        self.publicKey = publicKey
        self.privateKey = privateKey
        self.tocken = ""
        self.updateTime = time.time()
        self.getSuccess = False
        self.expiredTime = 0


    def GetMd5(self,timeMs): 
        strData = self.publicKey + self.pid + timeMs + self.privateKey
        checkcode = hashlib.md5(strData.encode()).hexdigest()
        print ("md5 = ",checkcode)
        return checkcode

    def GetJsonBody(self):
        mapRequest = {}
        mapRequest["productId"] = self.pid
        mapRequest["publicKey"] = self.publicKey
        timeMs = str(int(time.time() * 1000))
        mapRequest["sign"] = self.GetMd5(timeMs)
        mapRequest["timeStamp"] = timeMs
        return json.dumps(mapRequest)

    def OnTimer(self):
        curTime = time.time()
        if self.expiredTime == 0 or self.expiredTime - curTime < 60:
            self.Update()

    def Update(self):
        try:
            url = "https://api.talkinggenie.com/api/v2/public/authToken"
            payload = self.GetJsonBody()
            headers= {}
            headers["CONTENT-TYPE"] = "application/json;charset=UTF-8"  
            response = requests.request("POST", url, headers=headers, data = payload)
            res = response.text.encode('utf-8') 
            self.parseRequet(res)       
            print (res)
            return self.getSuccess
        except:
            self.getSuccess = False
            self.tocken = ""
            self.expiredTime = 0  
            strLog = "[ pid = " + self.pid + " publicKey = " + self.publicKey + " secretKey = " + self.privateKey + " time = " + time.strftime('%Y-%m-%d %H:%M:%S',time.localtime()) + "  update token fail]"
            log(strLog)
            print (strLog)
            return self.getSuccess

    def parseRequet(self,res):
        try:
            obj = json.loads(res)
            if obj["code"] != "200":
                self.getSuccess = False
                print ("get tocken fail ",res)
            else:
                self.tocken = obj["result"]["token"]
                self.expiredTime = int(obj["result"]["expireTime"])
                self.getSuccess = True
        except:
            self.getSuccess = False
            self.tocken = ""
            self.expiredTime = 0

    def GetTocken(self):
        return self.tocken


class CTask:
    def __init__(self,fileName,completeCallback,pid,hostName,uploadApi,getApi,separate,token):
        self.fileName = fileName
        self.PID = pid
        self.separate = separate
        self.status = TASK_STATUS_PRE
        self.fileId = ""
        self.fileUpCompleteTime = 0
        self.result = None
        self.completeCallback = completeCallback
        self.errorDes = "" 
        self.hostName = hostName
        self.uploadApi = uploadApi
        self.getApi = getApi
        self.retry = 0
        self.token = token

    def Run(self):
        while (self.status != TASK_STATUS_FIN and self.status != TASK_STATUS_FAIL):
            if not self.token:
                self.status = TASK_STATUS_FAIL
                self.errorDes = self.errorDes + " [has not token] "
                logContent = " fileName = " + self.fileName + " " + self.errorDes
                log(logContent)                

            if self.status == TASK_STATUS_PRE:
                self.sendRequest()
                time.sleep(1)


            if self.status == TASK_STATUS_RUN:
                if self.IsNeedDispatch():
                    self.getResult() 

        print ("task complete")
        self.completeCallback(self)

    def createRequestParameter(self):
        jsonParameter = {}
        dialog = {}
        parameter = {}
        dialog["productId"] = "12345678"
        metaObject = {}
        metaObject["recordId"] = "123456789"
        parameter["dialog"] = dialog
        parameter["metaObject"] = metaObject
        jsonParameter["param"] = parameter 
        if self.separate:
            data = {'param': '{"dialog": {"productId": "%s"},"metaObject": {"recordId": "%s","speechSeparate":true}}'%(self.PID,str(uuid.uuid1()))}
        else:
            data = {'param': '{"dialog": {"productId": "%s"},"metaObject": {"recordId": "%s","speechSeparate":false}}'%(self.PID,str(uuid.uuid1()))} 
        print (data)
        return data

    def parseRequet(self,res):
        try:
            error = res
            jsonRsult = json.loads(res)
            print (jsonRsult)
            if(jsonRsult["code"] != 200):
                self.status = TASK_STATUS_FAIL
                self.errorDes =  + self.errorDes + " [" + error +  "] " 
            else:
                self.fileId = jsonRsult["data"]["fileId"]
                self.status = TASK_STATUS_RUN 
                self.fileUpCompleteTime = time.time()
        except:
            if self.retry == RETRY_TIMES:
                self.status = TASK_STATUS_FAIL
            else:
                self.retry += 1
            self.errorDes = self.errorDes + "[ parse uploadFile response error ]"
            logContent = " fileName = " + self.fileName + " " + self.errorDes 
            log(logContent)

    def sendRequest(self):
        try: 
            url = self.hostName + self.uploadApi
            payload = self.createRequestParameter()  
            files = [
                ('file', open(self.fileName,'rb'))
            ]
            headers= {}
            headers["X-AISPEECH-TOKEN"] = self.token
            headers["X-AISPEECH-PRODUCT-ID"] = self.PID   
            #headers["Content-Type"] = "multipart/form-data"
            headers["Accept"] = "application/json"
            print(headers,payload)
            #print url,payload
            response = requests.request("POST", url, headers=headers, data = payload, files = files)
            res = response.text.encode('utf-8') 
            self.parseRequet(res)
        except:
            if self.retry == RETRY_TIMES:
                self.status = TASK_STATUS_FAIL
            else:
                self.retry += 1
            self.errorDes = self.errorDes + " [uploadFile error] "
            logContent = "data=" + str(payload) + " fileName = " + self.fileName + " " + self.errorDes
            log(logContent)

    def parseResult(self,res):
        try:
            error = res
            print (res)
            res = json.loads(res)
            if res["code"] != 200:
                self.status = TASK_STATUS_FAIL
                self.errorDes = self.errorDes + " [" + error + "] "
                return


            if res["data"]["status"] != "SUCCEED":
                self.status = TASK_STATUS_RUN 
                print ("waiting for result",self.fileName)

            else:
                self.status = TASK_STATUS_FIN 
            self.result = res  

        except Exception as e:
            print(e)
            self.status = TASK_STATUS_FAIL
            self.errorDes = self.errorDes + " [parse get result response error] "
            log(self.errorDes)


    def getResult(self):
        try:
            url = self.hostName + self.getApi
            if not self.fileId:
                log("fileId is null") 
            mapPid = {}
            mapPid["productId"] = self.PID
            mapObj = {}
            mapObj["fileId"] = self.fileId
            mapLoad = {}
            mapLoad["dialog"] = mapPid
            mapLoad["metaObject"] = mapObj
            payload = json.dumps(mapLoad)

            headers= {}
            headers["X-AISPEECH-TOKEN"] = self.token
            headers["X-AISPEECH-PRODUCT-ID"] = self.PID    
            headers["CONTENT-TYPE"] = "application/json;charset=UTF-8"  
            headers["Accept"] = "application/json;charset=UTF-8" 
            print(headers,payload)
            response = requests.request("POST", url, headers=headers, data = payload)
            res = response.text.encode('utf-8') 
            self.parseResult(res) 
        except:
            self.errorDes = self.errorDes + " [getResult error " + " fileId = " + self.fileId + "] "
            log(self.errorDes)


    def IsNeedDispatch(self):  
        if(time.time() - self.fileUpCompleteTime < 5):
            return False

        self.fileUpCompleteTime = time.time() 
        return True

    def IsTaskComplete(self):
        return self.status == TASK_STATUS_FAIL or self.status == TASK_STATUS_FIN

    def FormatResult(self):
        listResults = self.result["data"]["result"]
        formatResults = []
        for result in listResults:
            #formatResult = "beginTime = " + str(result["beginTime"]) + " | endTime = " + str(result["endTime"]) + " | channeId = " + str(result["channelId"]) + " | emotionValue = " + str(result["emotionValue"]) + " | speechRate = " + str(result["speechRate"]) + " | silenceDuration " + str(result["silenceDuration"]) + "         | text = " + result["text"] + "\n"
            formatResult = "%8d%15d%15d%20d%25d%25d          %s" % (result["beginTime"],result["endTime"],result["channelId"],result["emotionValue"],result["speechRate"],result["silenceDuration"],result["text"])            
            formatResult = formatResult + "\n"
            formatResults.append(formatResult)
        return formatResults


    def FormatResult2(self):
        listResults = self.result["data"]["result"]
        formatResults = []
        for result in listResults:
            #formatResult = "beginTime = " + str(result["beginTime"]) + " | endTime = " + str(result["endTime"]) + " | channeId = " + str(result["channelId"]) + " | emotionValue = " + str(result["emotionValue"]) + " | speechRate = " + str(result["speechRate"]) + " | silenceDuration " + str(result["silenceDuration"]) + "         | text = " + result["text"] + "\n"
            formatResult = result["text"]            
            formatResult = formatResult + "\n"
            formatResults.append(formatResult)
        return formatResults    

    def formatResult(self):
        try:
            if self.status == TASK_STATUS_FAIL:
                fileName = self.fileName.split(".")[0] + ".error"
                hFile = open(fileName,"wb")
                hFile.write(self.errorDes.encode())
                hFile.write(b"\n")
                if self.result:
                    hFile.write(self.result.encode("utf-8"))
                hFile.close()

            if self.status == TASK_STATUS_FIN:
                fileName = self.fileName.split(".")[0] + ".success"
                hFile = open(fileName,"wb") 
                header = "beginTime     "  +  "     endTime     " + "     channeId     "  +  "     emotionValue     "  + "      speechRate     " + "     silenceDuration     " + "     text     " + "\n"
                hFile.write(header.encode())
                formatResults = self.FormatResult() 
                for line in formatResults:
                    hFile.write(line.encode("utf-8")) 
                hFile.close()

                fileName = self.fileName.split(".")[0] + ".successNum"
                hFile = open(fileName,"wb") 
                formatResults = self.FormatResult2() 
                for line in formatResults:
                    hFile.write(line.encode("utf-8")) 
                hFile.close()                        

        except Exception as e:
            print ("format file error",e)
            fileName = self.fileName.split(".")[0] + ".error"
            hFile = open(fileName,"w")
            hFile.write("format file error\n")
            hFile.write(traceback.format_exc())
            hFile.close()            


class CTaskManager:
    def __init__(self,taskPath,workerNum,pid,hostName,uploadApi,getApi,publicKey,secretKey,separate=True):
        self.workerNum = workerNum
        self.separate = separate
        self.taskPath = taskPath
        self.workers = []
        self.taskQueue = queue.Queue()
        self.completeQueue = queue.Queue()
        self.totalTask = 0
        self.totalComplete = 0
        self.pid = pid
        self.hostName = hostName
        self.resultDes = ""
        self.uploadApi = uploadApi
        self.getApi = getApi
        self.publicKey = publicKey
        self.secretKey = secretKey
        self.auth = CAuthentication(self.pid,self.publicKey,self.secretKey)

    def getFiles(self):
        try:
            listFile = []
            for f in os.listdir(self.taskPath):
                if os.path.isdir(f):
                    continue
                if f.find("py") != -1 or f.find(".error") != -1 or f.find(".success") != -1:
                    continue
                f = self.taskPath + "/" + f
                listFile.append(f)  
            return listFile     
        except:
            pass

    def Init(self):
        for i in range(self.workerNum):
            worker = CWorkerThread(i,"thread"+str(i),self.GetTask,self.IsTaskComplete)
            worker.setDaemon(True)
            self.workers.append(worker)

        listFiles = self.getFiles()
        if not listFiles:
            return

        for fileName in listFiles:
            task = CTask(fileName,self.CompleteCallback,self.pid,self.hostName,self.uploadApi,self.getApi,self.separate,self.auth.GetTocken()) 
            self.taskQueue.put(task)
            self.totalTask = self.taskQueue.qsize()

        self.totalComplete = 0

    def Start(self):
        for worker in self.workers:
            worker.start()

    def GetTask(self,timeout = 1):
        return self.taskQueue.get(timeout=1)

    def InsertQueue(self,Task):
        self.taskQueue.put(Task)

    def CompleteCallback(self,Task):
        if not Task.IsTaskComplete():
            self.InsertQueue(Task) 
        else:
            print ("task complete name = ",Task.fileName,"  code = ",Task.status)
            Task.formatResult()
            self.InserCompleteQueue(Task)
            self.totalComplete = self.totalComplete + 1

    def InserCompleteQueue(self,Task): 
        self.completeQueue.put(Task)


    def IsTaskComplete(self):
        if(self.totalComplete == self.totalTask):
            return True
        else:
            return False 

    def Join(self):
        while True:
            if self.IsTaskComplete():
                return
            time.sleep(2)

    def formatResult(self):
        while self.completeQueue.qsize() > 0:
            task = self.completeQueue.get(timeout = 1)
            task.formatResult()

    def GetResultDes(self):
        success = 0
        fail = 0
        while self.completeQueue.qsize() > 0:
            task = self.completeQueue.get(timeout = 1)
            if task.status == TASK_STATUS_FIN:
                success = success + 1
            else:
                fail = fail + 1

        self.resultDes = "totalFile = " + str(self.totalTask) +  " : success = " + str(success) + ", fail = " + str(fail)
        return self.resultDes

    def getProcess(self):
        if self.totalComplete == self.totalTask:
            return "totalFile = " + str(self.totalTask) + "   process............  %100"


        fRatio = float(self.totalComplete) / self.totalTask * 100.0
        fRatio = "totalFile = " + str(self.totalTask) + "    process..........%" + "%4.2f"%fRatio
        return fRatio

    def OnTimer(self):
        if self.auth:
            self.auth.OnTimer()

    def IsTokenOk(self):
        try:
            if not self.auth:
                return False

            return self.auth.Update()
        except:
            return False

class numberInvThr(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        wb = Workbook()
        sheet = wb.active  
        sheet.title = con_excelSheetName
        sheet['A1'] = u'录音'
        sheet['B1'] = u'city_code'
        sheet['C1'] = u'转写内容（阿拉伯数字）'
        sheet['D1'] = u'转写内容（中文数字）'
        sheet['E1'] = u'地址识别'
        sheet['F1'] = u'地址补全'
        sheet['G1'] = u'警情要素'


        global g_isNumberInverting,g_finNum
        l = os.listdir(g_filePath)
        writePos = 2
        for one in l:
            if not '.successNum' in one:
                continue
            print (one)
            oneFile = g_filePath+'/'+one
            f = open(oneFile,'r',encoding='utf8')
            li = f.readlines()
            f.close()
            recordName = one.split('.successNum')[0]
            textChinese = u''
            textAlabo = u''
            for oo in li:
                textChinese += oo
                textAlabo += numChg(oo) 
                g_finNum += 1

            sheet['A%s'%writePos] = recordName
            sheet['C%s'%writePos] = textAlabo
            sheet['D%s'%writePos] = textChinese

            writePos += 1
        wb.save(g_filePath+'/'+con_excelName)  
        g_isNumberInverting = False
        print ('change over')
        
class CWorkerThread(threading.Thread):   
    def __init__(self, threadID,name,GetTask,IsTaskComplete):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
        self.GetTask = GetTask
        self.IsTaskComplete = IsTaskComplete

    def IsTaskOver(self):
        return self.IsTaskComplete()

    def run(self):                  
        while True:
            try:
                if(self.IsTaskOver()):
                    break

                time.sleep(1)
                task = self.GetTask()
                if task:
                    task.Run()
            except:
                break

        print ("thread exit:",self.getName())



class MyFrame(wx.Frame):
    def __init__(self):
        self.hasNewTask = False
        wx.Frame.__init__(self,None, -1, title=u'批量质检工具',size=(500, 550)) 
        self.panel = wx.Panel(self, size=(1500, 1200)) 
        self.Pidtext=wx.StaticText(self.panel,label='productid:',pos = (10,10))
        self.PidEditText = wx.TextCtrl(self.panel, value = con_productId,pos = (110,10),size=(100,20))

        self.publictext=wx.StaticText(self.panel,label='publicKey:',pos = (10,40))
        self.publicEditText = wx.TextCtrl(self.panel, value = con_pubKey,size = (300,25),pos = (110,40))
        self.secrettext=wx.StaticText(self.panel,label='secretKey:',pos = (10,80))
        self.secretEditText = wx.TextCtrl(self.panel, value = con_secKey,size = (300,25),pos = (110,80))        

        self.workertext=wx.StaticText(self.panel,label='并发数:',pos = (10,120))
        self.wokerEditText = wx.TextCtrl(self.panel, value = "2",pos = (110,120))      

        self.separetetext=wx.StaticText(self.panel,label='话者分离:',pos = (240,120))
        #self.separeteEditText = wx.TextCtrl(self.panel, value = "false",pos = (300,40))    
        self.list1 = ["true","false"]
        self.ch1=wx.ComboBox(self.panel,-1,value=con_fenli,choices=self.list1,pos = (300,120),size=(100,20))
        self.Bind(wx.EVT_COMBOBOX,self.on_combox,self.ch1)        

        self.hostNameStatic=wx.StaticText(self.panel,label='hostName:',pos = (10,160))
        self.hostNameStaticEdit = wx.TextCtrl(self.panel, value = con_hostName,size = (300,25) ,pos = (110,160))    

        self.UploadAPIName=wx.StaticText(self.panel,label='upload api:',pos = (10,200))
        self.UploadAPINameStaticEdit = wx.TextCtrl(self.panel, value =con_upApi,size = (300,25) ,pos = (110,200))           

        self.GetAPIName=wx.StaticText(self.panel,label='get api:',pos = (10,240))
        self.GetAPINameStaticEdit = wx.TextCtrl(self.panel, value = con_getApi,size = (300,25) ,pos = (110,240)) 

        self.niLabel=wx.StaticText(self.panel,label='逆文本:',pos = (10,280))
        self.niStaticEdit = wx.TextCtrl(self.panel, value = con_niUrl,size = (300,25) ,pos = (110,280)) 
        
        
        self.bChooseDir = wx.Button(self.panel,-1,u"文件夹选择",pos=(10,320))
        self.Bind(wx.EVT_BUTTON, self.OnButton, self.bChooseDir) 
        self.showText = wx.StaticText(self.panel,label='',pos = (110,320),size = (300,25))

        self.status = wx.StaticText(self.panel,label='',pos = (10,360))
        self.status.SetBackgroundColour("Green")
        self.resultPos = wx.StaticText(self.panel, pos = (10,380))
        self.resultPos.SetBackgroundColour("Green")
        self.resultDes = wx.StaticText(self.panel, pos = (10,380))
        self.resultDes.SetBackgroundColour("Green")
        self.numberInvertLabel = wx.StaticText(self.panel, pos = (10,400))

        self.btn_hello = wx.Button(self.panel, label=u'begin',pos = (10,460))
        self.btn_exit = wx.Button(self.panel, label='exit',pos = (150,460))

        self.btn_hello.Bind(wx.EVT_BUTTON, self.on_taskbegin)
        self.btn_exit.Bind(wx.EVT_BUTTON, self.on_exit)

        # Fit方法使框架自适应内部控件
        #self.Fit()
        self.timer = wx.Timer(self) #创建定时器
        self.Bind(wx.EVT_TIMER, self.OnTimer, self.timer)#绑定一个定时器事件        
        self.timer.Start(1000)
        self.taskManager = None
        self.path = None
        self.workers = 2
        self.ProductId = "" 
        self.separate = False
        self.hostName = ""
        self.getApi = ""
        self.uploadApi = ""
        self.taskBegin = False
        self.publicKey = ""
        self.secretKey = ""

    def on_combox(self,event): 
        if event.GetString().find("true") != -1:
            self.separate = True
        else:
            self.separate = False


    def OnButton(self,env):
        dlg = wx.DirDialog(self,u"选择文件夹",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPath()  
            #self.showText.label = self.path
            self.showText.SetLabelText(self.path)
            #print dir(self.showText)
        dlg.Destroy()        


    def GetValue(self):
        try:
            wokers = self.wokerEditText.GetValue()
            if wokers:
                self.workers = int(wokers)
                print ("workers = ",self.workers)
        except:
            pass

        try:
            if self.separate:
                print ("separate = true")
            else:
                print ("separate = false" )
        except:
            pass        

        try:
            pid = self.PidEditText.GetValue()
            if pid:
                self.ProductId = pid
                print( "ProductId = ",self.ProductId)
        except:
            pass


        try:
            baseURL = self.hostNameStaticEdit.GetValue()
            if baseURL:
                self.hostName = baseURL
                print ("hostName = ",self.hostName)
        except:
            pass

        try:
            api = self.UploadAPINameStaticEdit.GetValue()
            if api:
                self.uploadApi = api
                print ("upload api = ",self.uploadApi)
        except:
            pass      

        try:
            api = self.GetAPINameStaticEdit.GetValue()
            if api:
                self.getApi = api
                print ("get api = ",self.getApi)
        except:
            pass          

        try:
            publiKey = self.publicEditText.GetValue()
            if publiKey:
                self.publicKey = publiKey
                print ("get public key = ",self.publicKey)
        except:
            pass

        try:
            secretKey = self.secretEditText.GetValue()
            if secretKey:
                self.secretKey = secretKey
                print ("get secret key = ",self.secretKey)
        except:
            pass

    def IsTokenContextOk(self):
        if not self.publicKey or not self.secretKey:
            return False

        return True


    def GetToken(self):
        if not self.IsTokenContextOk():
            self.status.SetLabelText(u"请输入公钥和密钥")
            return False

        return True

    def on_taskbegin(self, event):  
        self.GetValue() 
        self.resetShow()
        if not self.GetToken():
            return 

        if not self.path:
            self.status.SetLabelText(u"请选择文件夹")
            return
        global g_filePath,g_niUrl
        g_niUrl = self.niStaticEdit.GetValue()
        g_filePath = self.path        
        self.taskManager = CTaskManager(self.path,self.workers,self.ProductId,self.hostName,self.uploadApi,self.getApi,self.publicKey,self.secretKey,self.separate)
        if not self.taskManager.IsTokenOk():
            self.status.SetLabelText(u"输入的公钥或密钥有误请重新输入")
            return

        self.taskManager.Init() 
        self.taskManager.Start()    
        self.taskBegin = True 

    def on_exit(self, evt):
        """退出程序"""
        wx.Exit()

    def resetShow(self):
        self.status.SetLabelText("")  
        self.resultPos.SetLabelText("")  
        self.resultDes.SetLabelText("") 
        self.numberInvertLabel.SetLabelText("") 

    def OnTimer(self,evt): 
        global g_isNumberInverting,g_finNum
        
        if self.hasNewTask and not g_isNumberInverting :
            self.hasNewTask = False
            self.numberInvertLabel.SetLabelText(u'数字转换已完成')       
            
        if self.hasNewTask and  g_isNumberInverting :
            self.numberInvertLabel.SetLabelText(u'正在进行数字转换,完成行数%s'%g_finNum)       
            
        if not self.taskManager:
            return

        if not self.taskBegin:
            return

        self.taskManager.OnTimer()
        ratio = self.taskManager.getProcess() 
        self.status.SetLabelText(ratio)

        if(self.taskManager.IsTaskComplete()):
            reslutDes = self.taskManager.GetResultDes()
            self.status.SetLabelText("tasks complete : " + reslutDes)  
            self.resultPos.SetLabelText(u"结果在目录下: " + self.path)  
            self.resultDes.SetLabelText(u"注：.error表示质检失败,.success表示质检成功") 
            g_finNum = 0
            self.numberInvertLabel.SetLabelText(u'正在进行数字转换,完成行数%s'%g_finNum)
            self.taskBegin = False
            g_isNumberInverting = True
            self.hasNewTask = True
            self.xx = t = numberInvThr()
            t.setDaemon(True)
            t.start()



if __name__ == "__main__": 
    #g_filePath = 'data'
    #t = numberInvThr()
    #t.setDaemon(True)
    #t.start()
    #while True:
        #time.sleep(1)    
        
    app = wx.App()
    myframe = MyFrame()
    myframe.Show()
    app.MainLoop()    

