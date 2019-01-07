#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from YouTubeAPI import YouTubeSearch as Search,YouTubeVideo_by_vId as Video,YouTubeChannel_by_cId as Channel
import json as js
import xlwings as xw
import time
import os


# In[ ]:


ResultPath="查找紀錄/"
RecordPath="查找結果/"
FileName="查找結果統計_{}.xlsx"
cURL="https://www.youtube.com/channel/"
vURL="https://www.youtube.com/watch?v="
CreateTime=time.strftime("%y%m%d%H%M",time.localtime())
nC=0
nV=0
WB=""
WSCI=""
WSCD=""
WSVI=""
WSVD=""


# In[ ]:


def SetExcel(Keyword):
    global WB,WSCI,WSCD,WSVI,WSVD,FileName
    FileName=FileName.format(Keyword)
    if os.path.isfile(FileName):
        WB=xw.Book(FileName)
        
    else:
        WB=xw.Book('查找結果統計.xlsx')
        WB.save(FileName)
    WSCI=WB.sheets["頻道資訊"]
    WSCD=WB.sheets["頻道計算"]
    WSVI=WB.sheets["影片資訊"]
    WSVD=WB.sheets["影片計算"]


# In[ ]:


# Result=Search("LNG")[1]
# Result


# In[ ]:


def SearchResult(Result):
    Data={}
    for SR in Result:
        cId=SR["snippet"]["channelId"]
        cNm=SR["snippet"]["channelTitle"]
        vId=SR["id"]["videoId"]
        vNm=SR["snippet"]["title"]
        if(cId not in Data):Data[cId]=[]
        Data[cId].append(vId)
    return Data


# In[ ]:


# SR=SearchResult(Result)
# SR=js.dumps(SR, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)


# In[ ]:


def SearchDetail(SR):
    SD={}
    #CreateTime=time.strftime("%y%m%d%H%M",time.localtime())
    for cId in SR:
        cD=Channel(cId)
        if cD==[]:continue
        cD=cD[0]
        cInfo=cD["snippet"]
        if "country" not in cInfo:
#             print("-"*50)
#             print(cInfo)
#             print("-"*50)
            continue
#         print(cInfo["country"])
        if cInfo["country"]!="TW":continue
        cNm=cInfo["title"]
        SD[cNm]={}
        cData=SD[cNm]
        cData["channelTitle"]=cNm
        cData["channelId"]=cId
        cData["channelUrl"]=cURL+cId
        if "customUrl" in cInfo:cData["customID"]=cInfo["customUrl"]
        cData["description"]=cInfo["description"]
        cData["publishedAt"]=cInfo["publishedAt"]
        cData["cCount"]={}
        cData["cCount"][CreateTime]=cD["statistics"]
        cData["Video"]={}
        for vId in SR[cId]:
            vD=Video(vId)[0]
            vInfo=vD["snippet"]
            vNm=vInfo["title"]
            cData["Video"][vNm]={}
            vData=cData["Video"][vNm]
            vData["videoTitle"]=vNm
            vData["videoId"]=vId
            vData["videoUrl"]=vURL+vId
            vData["description"]=vInfo["description"]
            vData["publishedAt"]=vInfo["publishedAt"]
            vData["tags"]=vInfo["tags"] if ("tags" in vInfo) else "Null"
            vData["vCount"]={}
            vData["vCount"][CreateTime]=vD["statistics"]
    return SD


# In[ ]:


# ChannelData=Channel("UCiw1PpoyhQ-CbgIgG_2_Jsg")
# ChannelData


# In[ ]:


# SD=SearchDetail(SR)


# In[ ]:


def OutputResultFile(Keyword,SD,max):
    global ResultPath
    #CreateTime=time.strftime("%y%m%d%H%M",time.localtime())
    jsFile=CreateTime+"_"+Keyword+"_"+str(max)+".json"
    with open(ResultPath+jsFile,"w", encoding="utf-8") as fp:
        js.dump(SD,fp, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)


# In[ ]:


def MergeRecord(SD,OR={}):
    global nC,nV
    for cNm in SD:
        cId=SD[cNm]["channelId"]
        if cNm not in OR:
            OR[cNm]=SD[cNm]
            nC+=1
        AddCannelInfo(WSCI,OR[cNm]) #Excel
        for Record in OR[cNm]["cCount"]:
            if Record not in OR[cNm]["cCount"]:
                OR[cNm]["cCount"][Record]=SD[cNm]["cCount"][Record]
            AddChannelCount(WSCD,cId,OR[cNm]["cCount"][Record],Record) #Excel
        for vNm in SD[cNm]["Video"]:
            vId=SD[cNm]["Video"][vNm]["videoId"]
            if vNm not in OR[cNm]["Video"]:
                OR[cNm]["Video"][vNm]=SD[cNm]["Video"][vNm]
                nV+=1
            AddVideoInfo(WSVI,cNm,OR[cNm]["Video"][vNm]) #Excel
            for Record in SD[cNm]["Video"][vNm]["vCount"]:
                if Record not in OR[cNm]["Video"][vNm]["vCount"]:
                    OR[cNm]["Video"][vNm]["vCount"][Record]=SD[cNm]["Video"][vNm]["vCount"][Record]
                AddVideoCount(WSVD,vId,OR[cNm]["Video"][vNm]["vCount"][Record],Record)
    WB.save()
    return OR


# In[ ]:


def OutputRecordFile(Keyword,SD):
    global nC,nV
    OR={}
    nC=0
    nV=0
    for file in os.listdir(RecordPath):
        if file.endswith(".json"):
            fileNm=file[:-5] #去除 ".json"
            fileNm=fileNm.split("_")
            if fileNm[0]==Keyword:
                nC=int(fileNm[1])
                nV=int(fileNm[2])
                jsFile=file
                with open(RecordPath+file , 'r', encoding='UTF-8') as jf:
                    OR = js.loads(jf.read()) #Orginal Record
                break
    else:
        jsFile="{}_{}_{}.json".format(Keyword,nC,nV)
    SetExcel(Keyword)
    Result=MergeRecord(SD,OR)
    
    with open(RecordPath+jsFile,"w", encoding="utf-8") as fp:
        js.dump(Result,fp, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)
        
    os.rename(RecordPath+jsFile,RecordPath+"{}_{}_{}.json".format(Keyword,nC,nV))


# In[ ]:


def AddCannelInfo(WS,cData):
    iM=RowMax(WS)+1
    if FindPT(WS.cells.columns[0],cData["channelId"])==None:
        WS.cells(iM,1).value=cData["channelId"]
        if "customID" in cData:WS.cells(iM,2).value=cData["customID"]
        WS.cells(iM,3).value=cData["channelTitle"]#add_hyperlink(cData["channelUrl"],cData["channelTitle"])
        WS.cells(iM,4).value=cData["publishedAt"]
        cCount=cData["cCount"]
        for Record in cCount:
            WS.cells(iM,5).value=cCount[Record]["viewCount"]
            WS.cells(iM,6).value=cCount[Record]["videoCount"]
            WS.cells(iM,7).value=cCount[Record]["subscriberCount"]
            WS.cells(iM,8).value=Record
        WS.cells(iM,9).value=cData["description"]
#     print("{} {}".format(WS.name,iM))


# In[ ]:


def AddChannelCount(WS,cId,cCount,Record):
    iM=RowMax(WS,2)+1
    if FindPT(WS.cells.columns[0],cId+Record)==None:
        WS.cells(iM,1).value=cId+Record
        WS.cells(iM,2).value=cId
        WS.cells(iM,3).value=Record
        WS.cells(iM,4).value=cCount["viewCount"]
        WS.cells(iM,5).value=cCount["videoCount"]
        WS.cells(iM,6).value=cCount["subscriberCount"]
#     print("  {} {}".format(WS.name,iM))


# In[ ]:


def AddVideoInfo(WS,cNm,vData):
    iM=RowMax(WS)+1
    if FindPT(WS.cells.columns[1],vData["videoId"])==None:
        WS.cells(iM,1).value=cNm
        WS.cells(iM,2).value=vData["videoId"]
        WS.cells(iM,3).value=vData["videoTitle"] #add_hyperlink(vData["videoUrl"],vData["videoTitle"])
        WS.cells(iM,4).value=vData["publishedAt"]
        vCount=vData["vCount"]
        for Record in vCount:
            WS.cells(iM,5).value=vCount[Record]["viewCount"]
            if "likeCount" in vCount[Record]:WS.cells(iM,6).value=vCount[Record]["likeCount"]
            if "dislikeCount" in vCount[Record]:WS.cells(iM,7).value=vCount[Record]["dislikeCount"]
            if "commentCount" in vCount[Record]:WS.cells(iM,8).value=vCount[Record]["commentCount"]
            WS.cells(iM,9).value=Record
        WS.cells(iM,10).value=vData["description"]
        strTag=""
        for Tag in vData["tags"]:
            if strTag=="":
                strTag=Tag
            else:
                strTag+=(","+Tag)
        WS.cells(iM,11).value=strTag
#     print("  {} {}".format(WS.name,iM))


# In[ ]:


def AddVideoCount(WS,vId,vCount,Record):
    iM=RowMax(WS,2)+1
    if FindPT(WS.cells.columns[0],vId+Record)==None:
        WS.cells(iM,1).value=vId+Record
        WS.cells(iM,2).value=vId
        WS.cells(iM,3).value=Record
        WS.cells(iM,4).value=vCount["viewCount"]
        if "likeCount" in vCount: WS.cells(iM,5).value=vCount["likeCount"]
        if "dislikeCount" in vCount: WS.cells(iM,6).value=vCount["dislikeCount"]
        if "commentCount" in vCount:WS.cells(iM,7).value=vCount["commentCount"]
#     print("    {} {}".format(WS.name,iM))


# In[ ]:


def FindPT(FindRng,What):
    PT=FindRng.api.cells.Find(What)
    return PT


# In[ ]:


def ColMax(WS,R=1,Colmin=1):
    C=WS.cells(R,WS.cells.columns.count).end("left").column
    C=max(C,Colmin)
    return C


# In[ ]:


def RowMax(WS,C=1,Rowmin=1):
    R=WS.cells(WS.cells.rows.count, C).end("up").row
    R=max(R,Rowmin)
    return R


# In[ ]:


def YouTubeSearch_by_Keyword(Keyword=None,max=50,token=None):
    CreateTime=time.strftime("%y%m%d%H%M",time.localtime())
    Result=Search(Keyword,max,token)
    SR=SearchResult(Result[1])
    SD=SearchDetail(SR)
#    OutputResultFile(Keyword,SD,max)
    OutputRecordFile(Keyword,SD)
    return Result[0]


# In[ ]:


# SD=SearchDetail(SearchResult(Search("彩妝")[1]))


# In[ ]:


# SetExcel("彩妝")
# vNm="便宜好用24款開架彩妝 Drugstore Must Have | 沛莉 Peri"
# vId="hPXz6NKKYuE"
# AddVideoInfo(WSVI,"百變沛莉 Peri",SD["百變沛莉 Peri"]["Video"][vNm])


# In[ ]:




