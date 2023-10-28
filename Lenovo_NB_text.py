from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import requests
import logging
import random
from util import web_driver

ID_link = []
Money_link = []
href_line = []

# 定義要爬取的網址
url = "https://www.lenovo.com/us/en/search?fq=&text=Laptops&rows=20&sort=relevance&fsid=1"
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}

#開啟搜尋頁面
Lenovo_NB = web_driver()
Lenovo_NB.get(url)
Lenovo_NB.execute_script("document.body.style.zoom='50%'")
sleep(2)
soup_num = BeautifulSoup(Lenovo_NB.page_source,'html.parser')
num_t = soup_num.select("p.show > span.page")
num_n = soup_num.select("p.show > span.total")
# #頁面向下滾動 & 資料載入

# while num_n[0].text != num_t[0].text:
#     #向下滾動
#     Lenovo_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     sleep(2)
#     Lenovo_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     #按按鈕
#     button = Lenovo_NB.find_elements(By.CSS_SELECTOR,'button.more')
#     Lenovo_NB.execute_script("arguments[0].click();", button[0])
#     sleep(2)
#     soup_num = BeautifulSoup(Lenovo_NB.page_source,'html.parser')
#     num_t = soup_num.select("p.show > span.page")
#     Lenovo_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     sleep(2)

#抓連結
sleep(5)
soup = BeautifulSoup(Lenovo_NB.page_source,'html.parser')
Href = soup.select("li.product_item")
Lenovo_NB.quit()
    
#將名稱與商品網址寫入list
for link in Href:
    ID = link.select("div.product_name > a")
    Money = link.select("div.price_box > span.final-price")
    if len(ID) > 0 and len(Money) > 0:
        ID_link.append(ID[0].text)
        new_money = Money[0].text.split("$")[-1]
        Money_link.append(new_money)
        href_line.append(ID[0]['href'])
    else:
        pass

#分別進入各商品頁面
i = 0

for i in range(len(href_line)):
    delay = random.uniform(1.0, 5.0)
    sleep(delay)
    print("Lenovo_NB {}".format(i))
    data_url = href_line[i] 
    Lenovo_NB_data = requests.get(data_url + "#features", headers=my_header)
    sleep(2)
    A = ["Model Name","Official Price","Web Link","Brand"]
    B = [ID_link[i],Money_link[i],data_url,"Lenovo"]    
    E = ["Model Name","Official Price","Web Link","Brand"]
    if Lenovo_NB_data.status_code==200:
        L_NB_soup = BeautifulSoup(Lenovo_NB_data.text,'html.parser')
        sleep(4)    
        #抓取特徵名稱
        L_NB_data_n = L_NB_soup.select("div.system_specs_container > ul > li > div.title")
        #抓取特徵內容
        L_NB_data_d = L_NB_soup.select("div.system_specs_container > ul > li > p")
        #商品詳細網址            
        NB_deta_url = L_NB_soup.select("div.system_specs_top > a")
    G4 = ""
    G5 = ""
    Weight = []
    if len(NB_deta_url) > 0:
        Data = {}
        j = 0        
        for j in range(len(L_NB_data_n)):
            #判斷已抓取資料是否重複
            if L_NB_data_n[j].text in E:
                pass
            else:
                #篩選需求特徵 & 單位換算處理 & 資料切割
                if L_NB_data_n[j].text =="Processor" or L_NB_data_n[j].text =="Memory" or L_NB_data_n[j].text =="Camera":
                    A.append(L_NB_data_n[j].text)
                    E.append(L_NB_data_n[j].text)
                    B.append(L_NB_data_d[j].text)
                elif "Display" in L_NB_data_n[j].text:
                    if "Display" not in A:
                        A.append("Display")
                        E.append(L_NB_data_n[j].text)
                        B.append(L_NB_data_d[j].text)
                elif "Operating System" in L_NB_data_n[j].text:
                    if "Operating System" not in A:
                        A.append("Operating System")
                        E.append(L_NB_data_n[j].text)
                        B.append(L_NB_data_d[j].text)
                elif 'AC Adapter' in L_NB_data_n[j].text or 'Power Supply' in L_NB_data_n[j].text:
                    A.append("Power Supply")
                    E.append("Power Supply")
                    B.append(L_NB_data_d[j].text)    
                elif "Graphic" in L_NB_data_n[j].text :
                    if "Graphics Card" not in A:
                        A.append("Graphics Card")
                        E.append("Graphics Card")
                        B.append(L_NB_data_d[j].text)
                elif L_NB_data_n[j].text =="Storage":
                    A.append("Hard Drive")
                    E.append(L_NB_data_n[j].text)
                    B.append(L_NB_data_d[j].text)
                elif L_NB_data_n[j].text =="WWAN":
                    if " 4G " in str(L_NB_data_d[j].text):
                        G4 = "yes"
                    if " 5G " in str(L_NB_data_d[j].text):
                        G5 = "yes"
                elif "Slots" in L_NB_data_n[j].text or "Ports" in L_NB_data_n[j].text:
                    A.append("Ports & Slots")
                    E.append(L_NB_data_n[j].text)
                    B.append(L_NB_data_d[j].text)
                elif L_NB_data_n[j].text =="Audio":
                    A.append("Audio and Speakers")
                    E.append(L_NB_data_n[j].text)
                    B.append(L_NB_data_d[j].text)
                elif L_NB_data_n[j].text =="Fingerprint Reader" or str(L_NB_data_n[j].text) =="Fingerprint Reader":
                    A.append("FPR")
                    B.append("Yes")
                elif L_NB_data_n[j].text =="Battery":
                    A.append("Primary Battery")
                    E.append(L_NB_data_n[j].text)
                    B.append(L_NB_data_d[j].text)               
                elif "Dimensions (H x W x D)" in L_NB_data_n[j].text and "Dimensions (H x W x D)" not in E:                   
                    E.append("Dimensions (H x W x D)")
                    Dim = (L_NB_data_d[j].text).split("mm")
                    H = Dim[0].split("x")[-1].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].strip()
                    W = Dim[1].split("x")[-1]
                    D = Dim[2].split("x")[-1].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].strip()
                    E.append("Height")
                    E.append("Width")
                    E.append("Depth")
                    A.append("Height")
                    A.append("Width")
                    A.append("Depth")
                    B.append(H)
                    B.append(W)
                    B.append(D)
                elif "Weight" in L_NB_data_n[j].text:
                    E.append(L_NB_data_n[j].text)
                    Weight.append(L_NB_data_d[j].text)
                   
        #進入詳細資料網頁
        delay = random.uniform(0.5, 5.0)
        sleep(delay)

        Lenovo_NB_data_deta = web_driver()
        Lenovo_NB_data_deta.get("https://www.lenovo.com" + NB_deta_url[0]['href']+"#features")
        Lenovo_NB_data_deta.execute_script("document.body.style.zoom='50%'")
        Lenovo_NB_data_deta.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        #加載資料
        try:
            button = Lenovo_NB_data_deta.find_elements(By.CSS_SELECTOR,'button.collapse')
            Lenovo_NB_data_deta.execute_script("arguments[0].click();", button[0])
            sleep(2)
        except:
            pass
        
        sleep(2)
        L_NB_deta_soup = BeautifulSoup(Lenovo_NB_data_deta.page_source,'html.parser')
        NB_deta_n = L_NB_deta_soup.select("tr.item")
        #抓取FPR
        FPR_data = L_NB_deta_soup.select("div.feature-content-section > div.description > p > span")
        Lenovo_NB_data_deta.quit()
        
        #檢驗FPR
        if "FPR" not in A:
            FPR = "No"
            j = 0
            for j in range(len(FPR_data)):
                if "finger" in FPR_data[j].text:
                    FPR = "Yes"
            A.append("FPR")
            B.append(FPR)
        j = 0
        H=[]
        W=[]
        #篩選需求特徵 & 單位換算處理 & 資料切割
        for j in range(len(NB_deta_n)):
            NB_deta_Name = NB_deta_n[j].select("th")
            if NB_deta_Name[0].text in E:
                pass
            else:
                if NB_deta_Name[0].text == "Processor" or "Operating System" in NB_deta_Name[0].text or "Display" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Memory" or NB_deta_Name[0].text =="Camera" or "Graphic" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Storage" or "Slots" in NB_deta_Name[0].text or "Ports" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Audio" or NB_deta_Name[0].text =="Battery" or "Dimensions" in NB_deta_Name[0].text or "Weight" in NB_deta_Name[0].text:                        
                    NB_deta_d = NB_deta_n[j].select("p")
                    E.append(NB_deta_Name[0].text)
                    k = 0
                    D = []
                    for k in range(len(NB_deta_d)):
                        D.append(NB_deta_d[k].text)
                    D = "\n".join(D)
                    if "Graphic" in NB_deta_Name[0].text:
                        if "Graphics Card" not in A:
                            A.append("Graphics Card")   
                            B.append(D)
                    elif NB_deta_Name[0].text =="Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif "Display" in NB_deta_Name[0].text:
                        if "Display" not in A:
                            A.append("Display")
                            B.append(D)
                    elif "Operating System" in NB_deta_Name[0].text:
                        if "Operating System" not in A:
                            A.append("Operating System")
                            B.append(D)
                    elif "Slots" in NB_deta_Name[0].text or "Ports" in NB_deta_Name[0].text:                        
                        A.append("Ports & Slots")
                        B.append(D)
                    elif NB_deta_Name[0].text =="Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif NB_deta_Name[0].text =="Battery":
                        A.append("Primary Battery")
                        B.append(D)
                    elif "Dimensions" in NB_deta_Name[0].text:
                        Dim = D.split("/")
                        D_cut = 0
                        for D_cut in range(len(Dim)):
                            if len(Dim[D_cut].split("mm")) > 2:
                                if "Height" in E:
                                    pass
                                else:
                                    E.append("Height")
                                    E.append("Width")
                                    E.append("Depth")
                                    Dim_1 = Dim[D_cut].split("mm")
                                    if len(Dim_1[2]) < 2:
                                        Dim_1 = Dim[D_cut].split("x")
                                        H = Dim_1[0].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].strip()
                                        W = Dim_1[1].split("mm")[0].strip()
                                        De = Dim_1[2].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    else:
                                        H = Dim_1[0].split("x")[-1].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].strip()
                                        W = Dim_1[1].split("x")[-1].split("mm")[0].strip()
                                        De = Dim_1[2].split("x")[-1].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    A.append("Height")
                                    A.append("Width")
                                    A.append("Depth")
                                    B.append(H)
                                    B.append(W)
                                    B.append(De)
                            elif "mm" in Dim[D_cut] and "inches" in Dim[D_cut]:
                                Dim = D.split("(")
                                hwd = 0
                                for hwd in range(len(Dim)):
                                    if "mm" in Dim[hwd]:
                                        H = Dim[hwd].split(":")[-1].split("x")[0].strip()
                                        W = Dim[hwd].split(":")[-1].split("x")[1].strip()
                                        De = Dim[hwd].split(":")[-1].split("x")[2].split("as")[-1].strip()
                                        A.append("Height")
                                        A.append("Width")
                                        A.append("Depth")
                                        B.append(H)
                                        B.append(W)
                                        B.append(De)
                            elif "mm" in Dim[D_cut] and (len(Dim[D_cut].split("mm")[0]) > len(Dim[D_cut].split("mm")[1])):
                                    H = Dim[D_cut].split("(mm")[0].split("x")[0].strip()
                                    W = Dim[D_cut].split("(mm")[0].split("x")[1].strip()
                                    De = Dim[D_cut].split("(mm")[0].split("x")[2].strip()
                                    A.append("Height")
                                    A.append("Width")
                                    A.append("Depth")
                                    B.append(H)
                                    B.append(W)
                                    B.append(De)           
                    elif "Weight" in NB_deta_Name[0].text:
                        Weight.append(D)
                    else:
                        A.append(NB_deta_Name[0].text)
                        B.append(D)
                if "in the" in NB_deta_Name[0].text:
                    NB_deta_d = NB_deta_n[j].select("p")                        
                    k = 0
                    for k in range(len(NB_deta_d)):
                        if (("w" in NB_deta_d[k].text.lower() and "battery" not in NB_deta_d[k].text.lower()) or "adapter" in NB_deta_d[k].text.lower()) and "Power Supply" not in A:
                            A.append("Power Supply")
                            B.append(NB_deta_d[k].text)
                        if "battery" in NB_deta_d[k].text.lower() and "Primary Battery" not in A:
                            A.append("Primary Battery")
                            B.append(NB_deta_d[k].text) 
                        
                        
    else:
        #網頁寫法不同處理
        delay = random.uniform(0.5, 5.0)
        sleep(delay)

        Lenovo_NB_data = web_driver()
        Lenovo_NB_data.get(data_url)  
        Lenovo_NB_data.execute_script("document.body.style.zoom='50%'")
        Lenovo_NB_data.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        #加載資料
        try:
            button = Lenovo_NB_data.find_elements(By.CSS_SELECTOR,'button.collapse')
            Lenovo_NB_data.execute_script("arguments[0].click();", button[0])
            sleep(2)
        except:
            pass
        L_NB_data_soup = BeautifulSoup(Lenovo_NB_data.page_source,'html.parser')
        L_NB_data_ack = L_NB_data_soup.select("div.specs-inner > table > tbody > tr.item")        
        FPR_data = L_NB_data_soup.select("div.feature-content-section > div.description > p > span")
        Lenovo_NB_data.quit()

        #檢驗FPR
        FPR = "No"
        u = 0
        for u in range(len(FPR_data)):
            if "finger" in FPR_data[u].text:
                FPR = "Yes"
        A.append("FPR")
        B.append(FPR)
        
        
        H=[]
        W=[]
        # print(len(L_NB_data_ack))
        #篩選需求特徵 & 單位換算處理 & 資料切割
        for ack in L_NB_data_ack:
            L_NB_data_n = ack.select("th")
            
            if L_NB_data_n[0].text in E:
                pass
            else:                           
                if L_NB_data_n[0].text == "Processor" or "Operating System" in L_NB_data_n[0].text or "Display" in L_NB_data_n[0].text or L_NB_data_n[0].text =="Memory" or L_NB_data_n[0].text =="Camera" or "Graphic" in L_NB_data_n[0].text or L_NB_data_n[0].text =="Storage" or L_NB_data_n[0].text =="WLAN" or "Slots" in L_NB_data_n[0].text or "Ports" in L_NB_data_n[0].text or L_NB_data_n[0].text =="Audio" or "Battery" in L_NB_data_n[0].text or "Dimensions (H x W x D)" in L_NB_data_n[0].text or "Weight" in L_NB_data_n[0].text:
                    L_NB_data_d = ack.select("td > ul > li")
                    E.append(L_NB_data_n[0].text)
                    if len(L_NB_data_d) == 0:
                        L_NB_data_d = ack.select("td")
                    k = 0
                    D = []
                    for k in range(len(L_NB_data_d)):
                        D.append(L_NB_data_d[k].text)
                    D = "\n".join(D)
                    
                    if "Graphic" in L_NB_data_n[0].text:
                        if "Graphics Card" not in A:
                            A.append("Graphics Card")
                            B.append(D)
                    elif "Operating System" in L_NB_data_n[0].text:
                        if "Operating System" not in A:
                            A.append("Operating System")
                            B.append(D)
                    elif "Display" in L_NB_data_n[0].text:
                        if "Display" not in A:
                            A.append("Display")
                            B.append(D)
                    elif L_NB_data_n[0].text =="Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif L_NB_data_n[0].text =="WWAN":
                        if " 4G " in str(D):
                            G4 = "yes"
                        if " 5G " in str(D):
                            G5 = "yes"
                            
                    elif "Slots" in L_NB_data_n[0].text or "Ports" in L_NB_data_n[0].text:
                        A.append("Ports & Slots")
                        B.append(D)
                    elif L_NB_data_n[0].text =="Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif "Battery" in L_NB_data_n[0].text:
                        A.append("Primary Battery")
                        B.append(D)
                    elif "Dimensions (H x W x D)" in L_NB_data_n[0].text:
                        Dim = D.split("/")
                        D_cut = 0
                        for D_cut in range(len(Dim)):
                            if len(Dim[D_cut].split("mm")) > 2:
                                if "Height" in E:
                                    pass
                                else:
                                    E.append("Height")
                                    E.append("Width")
                                    E.append("Depth")
                                    Dim_1 = Dim[D_cut].split("mm")
                                    if len(Dim_1[2]) < 2:
                                        Dim_1 = Dim[D_cut].split("x")
                                        H = Dim_1[0].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].strip()
                                        W = Dim_1[1].split("mm")[0].strip()
                                        De = Dim_1[2].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    else:
                                        H = Dim_1[0].split("x")[-1].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].strip()
                                        W = Dim_1[1].split("x")[-1].split("mm")[0].strip()
                                        De = Dim_1[2].split("x")[-1].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    A.append("Height")
                                    A.append("Width")
                                    A.append("Depth")
                                    B.append(H)
                                    B.append(W)
                                    B.append(De)
                            elif "mm" in Dim and "inches" in Dim:
                                Dim = D.split("(")
                                hwd = 0
                                for hwd in range(len(Dim)):
                                    if "mm" in Dim[hwd]:
                                        H = Dim[hwd].split(":")[-1].split("x")[0].strip()
                                        W = Dim[hwd].split(":")[-1].split("x")[1].strip()
                                        De = Dim[hwd].split(":")[-1].split("x")[2].split("as")[-1].strip()
                                        A.append("Height")
                                        A.append("Width")
                                        A.append("Depth")
                                        B.append(H)
                                        B.append(W)
                                        B.append(De)
                                                                              
                    elif "Weight" in L_NB_data_n[0].text:
                        Weight.append(D)
                    else:
                        A.append(L_NB_data_n[0].text)
                        B.append(D)
                if "in the" in L_NB_data_n[0].text:
                    NB_deta_d = ack.select("td > ul > li > p")                       
                    k = 0
                    for k in range(len(NB_deta_d)):
                        if (("w" in NB_deta_d[k].text.lower() and "battery" not in NB_deta_d[k].text.lower()) or "adapter" in NB_deta_d[k].text.lower()) and "Power Supply" not in A:
                            A.append("Power Supply")
                            B.append(NB_deta_d[k].text)
                        if "battery" in NB_deta_d[k].text.lower() and "Primary Battery" not in A:
                            A.append("Primary Battery")
                            B.append(NB_deta_d[k].text)
    #極特殊處理案例 點選畫面與網路爬蟲開啟頁面不同
    if len(A) < 15:
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        Lenovo_NB_data_deta = web_driver()
        Lenovo_NB_data_deta.get(data_url+"#tech_specs")
  
        Lenovo_NB_data_deta.execute_script("document.body.style.zoom='50%'")
        Lenovo_NB_data_deta.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        #加載資料
        try:
            button = Lenovo_NB_data_deta.find_elements(By.CSS_SELECTOR,'button.collapse')
            Lenovo_NB_data_deta.execute_script("arguments[0].click();", button[0])
            sleep(2)
        except:
            pass
        
        sleep(2)
        L_NB_deta_soup = BeautifulSoup(Lenovo_NB_data_deta.page_source,'html.parser')
        Lenovo_NB_data_deta.quit()

        NB_deta_n = L_NB_deta_soup.select("tr.item")
        #抓取FPR
        FPR_data = L_NB_deta_soup.select("div.feature-content-section > div.description > p > span")
        
        #檢驗FPR
        if "FPR" not in A:
            FPR = "unknow"
            j = 0
            for j in range(len(FPR_data)):
                if "finger" in FPR_data[j].text:
                    FPR = "Yes"
            A.append("FPR")
            B.append(FPR)
        
        j = 0
        H=[]
        W=[]
        #篩選需求特徵 & 單位換算處理 & 資料切割
        for j in range(len(NB_deta_n)):
            NB_deta_Name = NB_deta_n[j].select("th > div")
            if NB_deta_Name[0].text in E:
                pass
            else:
                if NB_deta_Name[0].text == "Processor" or "Operating System" in NB_deta_Name[0].text or "Display" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Memory" or NB_deta_Name[0].text =="Camera" or "Graphic" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Storage" or "Slots" in NB_deta_Name[0].text or "Ports" in NB_deta_Name[0].text or NB_deta_Name[0].text =="Audio" or NB_deta_Name[0].text =="Battery" or "Dimensions" in NB_deta_Name[0].text or "Weight" in NB_deta_Name[0].text:
                    NB_deta_d = NB_deta_n[j].select("td > ul > li > p")
                    if len(NB_deta_d) <1:
                        NB_deta_d = NB_deta_n[j].select("td > p")
                    E.append(NB_deta_Name[0].text)
                    k = 0
                    D = []
                    for k in range(len(NB_deta_d)):
                        D.append(NB_deta_d[k].text)
                    D = "\n".join(D)
                    if "Graphic" in NB_deta_Name[0].text:
                        if "Graphics Card" not in A:
                            A.append("Graphics Card")   
                            B.append(D)
                    elif "Operating System" in NB_deta_Name[0].text:
                        if "Operating System" not in A:
                            A.append("Operating System")
                            B.append(D)
                    elif "Display" in NB_deta_Name[0].text:
                        if "Display" not in A:
                            A.append("Display")
                            B.append(D)
                    elif NB_deta_Name[0].text =="Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif "Slots" in NB_deta_Name[0].text or "Ports" in NB_deta_Name[0].text:                        
                        A.append("Ports & Slots")
                        B.append(D)
                    elif NB_deta_Name[0].text =="Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif NB_deta_Name[0].text =="Battery":
                        A.append("Primary Battery")
                        B.append(D) 
                    elif "Dimensions" in NB_deta_Name[0].text:
                        Dim = D.split("/")
                        D_cut = 0
                        for D_cut in range(len(Dim)):
                            if len(Dim[D_cut].split("mm")) > 2:
                                if "Height" in E:
                                    pass
                                else:
                                    E.append("Height")
                                    E.append("Width")
                                    E.append("Depth")
                                    Dim_1 = Dim[D_cut].split("mm")
                                    if len(Dim_1[2]) < 2:
                                        Dim_1 = Dim[D_cut].split("x")
                                        H = Dim_1[0].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].split("from")[-1].strip()
                                        W = Dim_1[1].split("mm")[0].strip()
                                        De = Dim_1[2].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    else:
                                        H = Dim_1[0].split("x")[-1].split(":")[-1].split("-")[-1].split("at")[-1].split("~")[-1].split("covers")[-1].split("chassis")[-1].split("as")[-1].split("–")[-1].split("aluminum")[-1].split("mm")[0].split("from")[-1].strip()
                                        W = Dim_1[1].split("x")[-1].split("mm")[0].strip()
                                        De = Dim_1[2].split("x")[-1].split("as")[-1].split("-")[-1].split("–")[-1].split("~")[-1].split("mm")[0].strip()
                                    A.append("Height")
                                    A.append("Width")
                                    A.append("Depth")
                                    B.append(H)
                                    B.append(W)
                                    B.append(De)
                            elif "mm" in Dim[D_cut] and "inches" in Dim[D_cut]:
                                Dim = D.split("(")
                                hwd = 0
                                for hwd in range(len(Dim)):
                                    if "mm" in Dim[hwd]:
                                        H = Dim[hwd].split(":")[-1].split("x")[0].split("from")[-1].strip()
                                        W = Dim[hwd].split(":")[-1].split("x")[1].strip()
                                        De = Dim[hwd].split(":")[-1].split("x")[2].split("as")[-1].split("mm")[0].strip()
                                        A.append("Height")
                                        A.append("Width")
                                        A.append("Depth")
                                        B.append(H)
                                        B.append(W)
                                        B.append(De)
                            elif "mm" in Dim[D_cut] and (len(Dim[D_cut].split("mm")[0]) > len(Dim[D_cut].split("mm")[1])):
                                    H = Dim[D_cut].split("(mm")[0].split("x")[0].split("from")[-1].strip()
                                    W = Dim[D_cut].split("(mm")[0].split("x")[1].strip()
                                    De = Dim[D_cut].split("(mm")[0].split("x")[2].split("mm")[0].strip()
                                    A.append("Height")
                                    A.append("Width")
                                    A.append("Depth")
                                    B.append(H)
                                    B.append(W)
                                    B.append(De)
                                
                    elif "Weight" in NB_deta_Name[0].text:
                        Weight.append(D)
                    else:
                        A.append(NB_deta_Name[0].text)
                        B.append(D)
                if "in the" in NB_deta_Name[0].text:
                    NB_deta_d = NB_deta_n[j].select("p")                        
                    k = 0
                    for k in range(len(NB_deta_d)):
                        if (("w" in NB_deta_d[k].text.lower() and "battery" not in NB_deta_d[k].text.lower()) or "adapter" in NB_deta_d[k].text.lower()):
                            if "Power Supply" not in A:
                                A.append("Power Supply")
                                B.append(NB_deta_d[k].text)
                        if "battery" in NB_deta_d[k].text.lower():
                            if "Primary Battery" not in A:
                                A.append("Primary Battery")
                                B.append(NB_deta_d[k].text)
    
    A.append("Wireless")
    if G4 != "" and G5 != "":
        B.append("4G/5G")
    elif G4 != "" and G5 == "":
        B.append("4G")
    elif G4 == "" and G5 != "":
        B.append("5G")
    else:
        B.append("No")
    #針對Weight處理 & 單位換算 & 資料切割
    symbol = ""                
    if len(Weight)>0:
        W_cut = Weight[0].split("at")[-1].split("from")[-1].split("than")[-1].split("Starting")[-1].split("g")
        if len(W_cut) > 1:
            if str(W_cut[0])[-1] == "K" or str(W_cut[0])[-1] == "k":
                # 
                if "<" in W_cut[0]: 
                    Wei = W_cut[0].split("K")[0].split("k")[0].split("/")[-1].split("(")[-1].split("<")[-1]
                    Wei = float(Wei)
                    symbol = "<"
                elif "Up" in W_cut[0]:
                    Wei = W_cut[0].split("K")[0].split("k")[0].split("/")[-1].split("(")[-1].split("to")[-1]
                    Wei = float(Wei)
                    symbol = ">"
                else:
                    Wei = W_cut[0].split("K")[0].split("k")[0].split("/")[-1].split("(")[-1].split("<")[-1]
                    Wei = float(Wei)
            else:
                if "<" in W_cut[0]: 
                    Wei = W_cut[0].split("/")[-1].split("(")[-1].split("<")[-1]
                    Wei = round(float(Wei)/1000,2)
                    symbol = "<"
                else:
                    Wei = W_cut[0].split("/")[-1].split("(")[-1].split("<")[-1]
                    Wei = round(float(Wei)/1000,2)
        else:
            Wei = ""
    else:
        Wei = ""

    A.append("Weight")
    B.append("{} {}".format(symbol,Wei))
    #資料合併
    A_dataframe = pd.Series(A)
    if i == 0:
        C_dataframe = pd.DataFrame(B,index = A_dataframe)
    else:
        B_dataframe = pd.DataFrame(B,index = A_dataframe)
        C_dataframe = C_dataframe.merge(B_dataframe,how = "outer",left_index=True, right_index=True)


C_dataframe.to_excel("Lenovo_NB.xlsx")