from apiclient.discovery import build
import json as js

key="AIzaSyBUaN5mX2k_jn744l8cq8fCfUQcOAEAkv4"
YouTubeAPI=build("youtube","v3",developerKey=key)
# key 必須使用 關鍵字參數 e.g. developerKey=key


# ------------------------------------------------------------------
def YouTubeSearch(q,max=50,token=None,SearchType=0,OrderType=0):  
    lstType=("video","channle","playlist")
    #影片、頻道、播放清單
    lstOrder=("relevance","rating","viewCount","videoCount","title","date")
    #關聯性、評價、觀看次數、影片數、標題、日期
    
    Result=YouTubeAPI.search().list(
        q=q,
        type=lstType[SearchType],
        order = lstOrder[OrderType],
        part="id,snippet",
        #part="id,snippet",  ##因為有videos/channels補足搜尋結果，故簡化搜尋結果。
        maxResults=max,
        pageToken=token
    ).execute()
    
    videos=[]
    for Element_Result in Result.get("items", []):
        if Element_Result["id"]["kind"] == "youtube#video":
            videos.append(Element_Result)
            
    try:
        nexttok = Result["nextPageToken"]
        return(nexttok, videos)
    except Exception as e:
        nexttok = "last_page"
        return(nexttok, videos)

# Result=YouTubeSearch("Leggy",max=1)
# Result=js.dumps(Result, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)
# print(Result)
# ------------------------------------------------------------------


# ------------------------------------------------------------------
def YouTubeVideo_by_vId(vId):
    Result=YouTubeAPI.videos().list(
        part="snippet,contentDetails,statistics",
        id=vId
    ).execute()

#     videos=[]
#     for Element_Result in Result.get("items", []):
#         if Element_Result["kind"] == "youtube#video":
#             videos.append(Element_Result)
            
    return Result.get("items", {})

# Result=YouTubeVideo_by_vId("-H1zciJDuvg")
# Result=js.dumps(Result, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)
# print(Result)
# ------------------------------------------------------------------


# ------------------------------------------------------------------
def YouTubeChannel_by_cId(cId):
    Result=YouTubeAPI.channels().list(
        part="snippet,statistics",
        id=cId
    ).execute()

#     channels=[]
#     for Element_Result in Result.get("items", []):
#         if Element_Result["kind"] == "youtube#channel":
#             channels.append(Element_Result)
            
    return Result.get("items", {})

# import json as js
# Result=YouTubeChannel_by_cId("UC2pAstZuQKWTYtY0TSkAmTQ")
# Result=js.dumps(Result, sort_keys=True, indent=4, separators=(',', ': '), ensure_ascii=False)
# print(Result)
# ------------------------------------------------------------------





