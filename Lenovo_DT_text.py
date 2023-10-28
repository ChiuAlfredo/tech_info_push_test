from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import requests
import logging
import random
from util import web_driver

# 定義要爬取的網址    
url = "https://www.lenovo.com/us/en/search?fq={!ex=prodCat}lengs_Product_facet_ProdCategories:PCs%20Tablets&text=Desktops&rows=20&sort=relevance&display_tab=Products"
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}
#開啟搜尋頁面
Lenovo_DT = web_driver()
Lenovo_DT.get(url)
Lenovo_DT.execute_script("document.body.style.zoom='50%'")
sleep(2)
soup_num = BeautifulSoup(Lenovo_DT.page_source,'html.parser')
num_t = soup_num.select("p.show > span.page")
num_n = soup_num.select("p.show > span.total")

# #頁面向下滾動 & 資料載入
# while num_n[0].text != num_t[0].text:        
#     #向下滾動
#     Lenovo_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     sleep(2)
#     Lenovo_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     #按按鈕
#     button = Lenovo_DT.find_elements(By.CSS_SELECTOR,'button.more')            
#     sleep(2)
#     Lenovo_DT.execute_script("arguments[0].click();", button[0])           
#     sleep(2)
#     soup_num = BeautifulSoup(Lenovo_DT.page_source,'html.parser')
#     num_t = soup_num.select("p.show > span.page")
#     Lenovo_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
   

#抓連結
sleep(5)
soup = BeautifulSoup(Lenovo_DT.page_source,'html.parser')
Href = soup.select("li.product_item")

Lenovo_DT.quit()

ID_link = []
Money_link = []
href_line = []
#將名稱與商品網址寫入list
for link in Href:
    ID = link.select("div.product_name > a")
    Money = link.select("div.price_box > span.final-price ")        
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
    print("Lenovo_DT {}".format(i))
    data_url = href_line[i]    
    Lenovo_DT_data = requests.get(data_url, headers=my_header)
    sleep(2)
    if Lenovo_DT_data.status_code==200:
        L_DT_soup = BeautifulSoup(Lenovo_DT_data.text,'html.parser')
        sleep(2)        
        #抓取特徵名稱
        L_DT_data_n = L_DT_soup.select("div.system_specs_container > ul > li > div.title")
        #抓取特徵內容
        L_DT_data_d = L_DT_soup.select("div.system_specs_container > ul > li > p")
        #商品詳細網址
        DT_deta_url = L_DT_soup.select("a.view-all-models")  
    A = ["Model Name","Official Price","Web Link","Brand"]
    E = ["Model Name","Official Price","Web Link","Brand"]
    B = [ID_link[i],Money_link[i],data_url,"Lenovo"]
    Dimensions = []
    H = ""
    W = ""
    De = ""
    if len(DT_deta_url) > 0:
        j = 0
        #判斷已抓取資料是否重複
        for j in range(len(L_DT_data_n)):
            if L_DT_data_n[j].text.replace(' ', '').strip() in E:
                pass
            else:
                DT_deta_Name = L_DT_data_n[j].text.replace(' ', '').strip()
                E.append(DT_deta_Name)
                #篩選需求特徵
                if DT_deta_Name == "Audio" or DT_deta_Name == "Display" or "Camera" in DT_deta_Name or "Graphic" in DT_deta_Name or DT_deta_Name == "Memory" or DT_deta_Name == "OperatingSystem" or DT_deta_Name == "Ports" or DT_deta_Name == "Ports/Slots" or DT_deta_Name == "Ports/Slots/Buttons" or DT_deta_Name == "Processor" or DT_deta_Name == "Storage" or DT_deta_Name == "WiFiWirelessLANAdapters":
                    k = 0
                    D = L_DT_data_d[j].text
                    if DT_deta_Name == "Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif "Camera" in DT_deta_Name:
                        A.append("Camera")
                        B.append(D)
                    elif DT_deta_Name == "Display":
                        A.append("Display")
                        B.append(D)
                    elif "Graphic" in DT_deta_Name:
                        if 'Graphics Card' not in A:
                            A.append("Graphics Card")
                            B.append(D)
                    elif DT_deta_Name == "Memory":
                        A.append("Memory")
                        B.append(D)
                    elif DT_deta_Name == "OperatingSystem":
                        A.append("Operating System")
                        B.append(D)
                    elif DT_deta_Name == "Ports" or DT_deta_Name == "Ports/Slots" or DT_deta_Name == "Ports/Slots/Buttons":
                        A.append("Ports & Slots")
                        B.append(D)
                    elif DT_deta_Name == "Processor":
                        A.append("Processor")
                        B.append(D)
                    elif DT_deta_Name == "Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif DT_deta_Name == "WiFiWirelessLANAdapters":
                        A.append("Wireless")
                        B.append(D)
        
        #進入詳細資料網頁
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        
        Lenovo_DT_data_deta = web_driver()          
        Lenovo_DT_data_deta.get("https://www.lenovo.com" + DT_deta_url[0]['href'] + "#features")
 
        Lenovo_DT_data_deta.execute_script("document.body.style.zoom='50%'")
        Lenovo_DT_data_deta.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        
        #加載資料
        try:
            button = Lenovo_DT_data_deta.find_elements(By.CSS_SELECTOR,'button.collapse')
            Lenovo_DT_data_deta.execute_script("arguments[0].click();", button[0])
            sleep(2)
        except:
            pass
        L_DT_deta_soup = BeautifulSoup(Lenovo_DT_data_deta.page_source,'html.parser')
        DT_deta_n = L_DT_deta_soup.select("tr.item")
        Lenovo_DT_data_deta.quit()

        #篩選需求特徵 & 單位換算處理 & 資料切割
        j = 0
        P ,R_Ports,F_Ports ="","",""
        for j in range(len(DT_deta_n)):
            DT_deta_Name = DT_deta_n[j].select("th")
            DT_deta_Name = DT_deta_Name[0].text.replace(' ', '').strip()
            if DT_deta_Name in E:
                pass
            else:
                if DT_deta_Name == "Audio" or DT_deta_Name == "Camera" or DT_deta_Name == "Camera/Mic" or "Dimension" in DT_deta_Name or "Graphic" in DT_deta_Name or DT_deta_Name == "Memory" or DT_deta_Name == "OperatingSystem" or "Ports" in DT_deta_Name or "Slots" in DT_deta_Name or DT_deta_Name == "Processor" or DT_deta_Name == "Storage" or "Weight" in DT_deta_Name or DT_deta_Name == "WiFiWirelessLANAdapters":
                    NB_deta_d = DT_deta_n[j].select("td")
                    k = 0
                    D = []
                    for k in range(len(NB_deta_d)):
                        D.append(NB_deta_d[k].text)
                    D = "\n".join(D)
                    if DT_deta_Name == "Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif DT_deta_Name == "Rear I/O Ports":
                        R_Ports = "Rear:" + D
                    elif DT_deta_Name == "Front I/O Ports":
                        F_Ports = "Front" + D
                    elif "Ports" in DT_deta_Name or "Slots" in DT_deta_Name:
                        P = P + "\n" + D
                    elif (DT_deta_Name == "Camera" or DT_deta_Name == "Camera/Mic") and ("Camera" not in A):
                        A.append("Camera")
                        B.append(D)
                    elif "Graphic" in DT_deta_Name:
                        if 'Graphics Card' not in A:
                            A.append("Graphics Card")
                            B.append(D)
                    elif DT_deta_Name == "Memory":
                        A.append("Memory")
                        B.append(D)
                    elif DT_deta_Name == "OperatingSystem":
                        A.append("Operating System")
                        B.append(D)
                    elif DT_deta_Name == "Ports" or DT_deta_Name == "Ports/Slots" or DT_deta_Name == "Ports/Slots/Buttons":
                        P = D
                    elif DT_deta_Name == "Processor":
                        A.append("Processor")
                        B.append(D)
                    elif DT_deta_Name == "Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif "Weight" in DT_deta_Name:
                        A.append("Weight")
                        if "kg" in D:
                            D = D.replace('Starting', ' ').split("kg")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = float(D)
                        elif "Kg" in D:
                            D = D.replace('Starting', ' ').split("Kg")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = float(D)
                        elif "g" in D:
                            D = D.replace('Starting', ' ').split("g")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = round(float(D)/1000,2)                               
                        B.append(D)
                    elif DT_deta_Name == "WiFiWirelessLANAdapters":
                        A.append("Wireless")
                        B.append(D)
                        
                    if "Dimension" in DT_deta_Name:
                        reD1 = D.split(":")[-1].split("/")
                        reD = 0
                        for reD in range(len(reD1)):
                            if "mm" in reD1[reD] and '"' not in reD1[reD] and "”" not in reD1[reD]:
                                N1 = reD1[reD].split("Stand:")
                                N1 = N1[-1].split("x")
                                N11 = N1[0].split("mm")[0].split("at")[-1].split("rom")[-1].split(":")[-1].strip()
                                N12 = N1[1].split("mm")[0].split("rom")[-1].split(":")[-1].strip()
                                N13 = N1[2].split("mm")[0].split("rom")[-1].split(":")[-1].strip()
                        if DT_deta_Name.split("(")[-1].split(")")[0].split("x")[0].strip()=="W" and DT_deta_Name.split("(")[-1].split(")")[0].split("x")[-1].strip()=="D":
                            D = "{}x{}x{}".format(N12,N11,N13)
                        elif DT_deta_Name.split("(")[-1].split(")")[0].split("x")[0].strip()=="W" and DT_deta_Name.split("(")[-1].split(")")[0].split("x")[-1].strip()=="H":
                            D = "{}x{}x{}".format(N13,N11,N12)
                        else:
                            D = "{}x{}x{}".format(N11,N12,N13)
                        Dimensions.append(D)

    else:
        delay = random.uniform(0.5, 5.0)
        sleep(delay)

        Lenovo_DT_data_deta = web_driver()          
        Lenovo_DT_data_deta.get(data_url + "#tech_specs")

        Lenovo_DT_data_deta.execute_script("document.body.style.zoom='50%'")
        Lenovo_DT_data_deta.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        #加載資料
        try:
            button = Lenovo_DT_data_deta.find_elements(By.CSS_SELECTOR,'button.collapse')
            Lenovo_DT_data_deta.execute_script("arguments[0].click();", button[0])
            sleep(2)
        except:
            pass
        L_DT_deta_soup = BeautifulSoup(Lenovo_DT_data_deta.page_source,'html.parser')
        DT_deta_n = L_DT_deta_soup.select("tr.item")
        Lenovo_DT_data_deta.quit()

        #篩選需求特徵 & 單位換算處理 & 資料切割
        j = 0
        P ,R_Ports,F_Ports ="","",""
        for j in range(len(DT_deta_n)):
            DT_deta_Name = DT_deta_n[j].select("th")
            DT_deta_Name = DT_deta_Name[0].text.replace(' ', '').strip()
            if DT_deta_Name in E:
                pass
            else:
                if DT_deta_Name == "Audio" or DT_deta_Name == "Camera" or DT_deta_Name == "Camera/Mic" or "Dimension" in DT_deta_Name or "Graphic" in DT_deta_Name or DT_deta_Name == "Memory" or DT_deta_Name == "OperatingSystem" or "Ports" in DT_deta_Name or "Slots" in DT_deta_Name or DT_deta_Name == "Processor" or DT_deta_Name == "Storage" or "Weight" in DT_deta_Name or DT_deta_Name == "WiFiWirelessLANAdapters":
                    NB_deta_d = DT_deta_n[j].select("td")
                    k = 0
                    D = []
                    for k in range(len(NB_deta_d)):
                        D.append(NB_deta_d[k].text)
                    D = "\n".join(D)
                    if DT_deta_Name == "Audio":
                        A.append("Audio and Speakers")
                        B.append(D)
                    elif DT_deta_Name == "Rear I/O Ports":
                        R_Ports = "Rear:" + D
                    elif DT_deta_Name == "Front I/O Ports":
                        F_Ports = "Front" + D
                    elif "Ports" in DT_deta_Name or "Slots" in DT_deta_Name:
                        P = P + "\n" + D
                    elif (DT_deta_Name == "Camera" or DT_deta_Name == "Camera/Mic") and ("Camera" not in A):
                        A.append("Camera")
                        B.append(D)
                    elif "Graphic" in DT_deta_Name:
                        A.append("Graphics Card")
                        B.append(D)
                    elif DT_deta_Name == "Memory":
                        A.append("Memory")
                        B.append(D)
                    elif DT_deta_Name == "OperatingSystem":
                        A.append("Operating System")
                        B.append(D)
                    elif DT_deta_Name == "Ports" or DT_deta_Name == "Ports/Slots" or DT_deta_Name == "Ports/Slots/Buttons":
                        P = D
                    elif DT_deta_Name == "Processor":
                        A.append("Processor")
                        B.append(D)
                    elif DT_deta_Name == "Storage":
                        A.append("Hard Drive")
                        B.append(D)
                    elif "Weight" in DT_deta_Name:
                        A.append("Weight")
                        if "kg" in D:
                            D = D.replace('Starting', ' ').split("kg")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = float(D)
                        elif "Kg" in D:
                            D = D.replace('Starting', ' ').split("Kg")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = float(D)
                        elif "g" in D:
                            D = D.replace('Starting', ' ').split("g")[0].split("(")[-1].split("at")[-1].split("rom")[-1].split("/")[-1].split(":")[-1].split("Around")[-1].strip()
                            D = round(float(D)/1000,2)                               
                        B.append(D)
                    elif DT_deta_Name == "WiFiWirelessLANAdapters":
                        A.append("Wireless")
                        B.append(D)
                        
                    if "Dimension" in DT_deta_Name:
                        reD1 = D.split(":")[-1].split("/")
                        reD = 0
                        for reD in range(len(reD1)):
                            if "mm" in reD1[reD] and '"' not in reD1[reD] and "”" not in reD1[reD]:
                                N1 = reD1[reD].split("Stand:")
                                N1 = N1[-1].split("x")
                                N11 = N1[0].split("mm")[0].split("at")[-1].split("rom")[-1].split(":")[-1].strip()
                                N12 = N1[1].split("mm")[0].split("rom")[-1].split(":")[-1].strip()
                                N13 = N1[2].split("mm")[0].split("rom")[-1].split(":")[-1].strip()
                        if DT_deta_Name.split("(")[-1].split(")")[0].split("x")[0].strip()=="W" and DT_deta_Name.split("(")[-1].split(")")[0].split("x")[-1].strip()=="D":
                            D = "{}x{}x{}".format(N12,N11,N13)
                        elif DT_deta_Name.split("(")[-1].split(")")[0].split("x")[0].strip()=="W" and DT_deta_Name.split("(")[-1].split(")")[0].split("x")[-1].strip()=="H":
                            D = "{}x{}x{}".format(N13,N11,N12)
                        else:
                            D = "{}x{}x{}".format(N11,N12,N13)
                        Dimensions.append(D)
    
    #針對 H/W/D 處理 單位換算處理 & 資料切割
    if len(Dimensions) > 0:
        H = Dimensions[0].split("x")[0].split("(")[0]
        W = Dimensions[0].split("x")[1]
        De = Dimensions[0].split("x")[2]
        
    A.append("Height")
    A.append("Width")
    A.append("Depth")
    B.append(H)
    B.append(W)
    B.append(De)
    
    #針對port與slot複合合併處理
    A.append("Ports & Slots")
    p_s =F_Ports +"\n"+ R_Ports +"\n"+ P
    B.append(p_s.strip())
    
    #資料合併
    A = pd.Series(A)
    if i == 0:
        C = pd.DataFrame(B,index = A)
    else:
        B = pd.DataFrame(B,index = A)
        C = C.merge(B,how = "outer",left_index=True, right_index=True)
        C.drop_duplicates(inplace=True)
       
C.to_excel("Lenovo_DT.xlsx")
