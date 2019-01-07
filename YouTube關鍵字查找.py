#!/usr/bin/env python
# coding: utf-8

# In[6]:


# 每次查找需要重新 Reset Kernel
# 輸入關鍵字
while True:
    Keyword=input("請輸入查找關鍵字：")
    if Keyword=="":
        print("關鍵字不得為空，請重新輸入。")
    else:
        break



# In[1]:


from YouTubeResult_v4 import YouTubeSearch_by_Keyword as Search


# In[2]:


print("\n本系統預設搜索500筆紀錄，並僅搜索台灣地區的頻道及影片。")
print("完成搜索後，系統會自動關閉。\n")
token=None
for i in range(10):
    token=Search(Keyword,token=token)
    print("{:>2}/10 {}".format(i+1,token))
    if token=="last_page":break


# In[ ]:




