from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import logging
import random
from util import web_driver

# 定義要爬取的網址
url = "https://www.hp.com/us-en/shop/sitesearch?keyword=Desktops"
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}

#開啟搜尋頁面
HP_DT = web_driver()
HP_DT.get(url)
HP_DT.execute_script("document.body.style.zoom='50%'")
sleep(2)

# #頁面向下滾動 & 資料載入

# while True:
#     try:
#         #向下滾動
#         HP_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         sleep(2)
#         HP_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         #按按鈕
#         button = HP_DT.find_elements(By.CSS_SELECTOR,'button.hawksearch-load-more')
#         HP_DT.execute_script("arguments[0].click();", button[0])
#         sleep(2)
#         HP_DT.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         sleep(2)
#     except:
#         break


sleep(5)
#抓連結
soup = BeautifulSoup(HP_DT.page_source,'html.parser')
Href = soup.select("div.ProductTile-module_wrapper__hS8-C.productTile")

HP_DT.quit()
ID_link = []
Money_link = []
href_line = []
href_line_new = []

#將名稱與商品網址寫入list
for link in Href:
    ID = link.select("div.ProductTile-module_details__Sgo-b > a > h3")
    Money = link.select("div.PriceBlock-module_salePriceWrapper__i7Rjp > div")
    href_line_new = link.select("div.ProductTile-module_details__Sgo-b > a")    
    ID_link.append(ID[0].text)
    if len(Money) > 0:
        new_money = Money[0].text.split("$")[-1]
        Money_link.append(new_money)
    else:
        Money_link.append("Nan")
    href_line.append("https://www.hp.com"+href_line_new[0]['href'])   

#分別進入各商品頁面
i = 0
for i in range(len(href_line)):
    print("HP_DT {}".format(i))
    delay = random.uniform(2.0,6.0)
    sleep(delay)
    data_url = href_line[i]
    HP_DT_data = web_driver()
    HP_DT_data.get(data_url)
    HP_DT_data.execute_script("document.body.style.zoom='50%'")
    HP_DT_data.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(4)
    L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
    L_DT_data_n = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
    L_DT_data_d_ack = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE")
    L_DT_specs = L_DT_soup.select("div.Container-module_root__luUPH.Container-module_container__jSUGk > div.Footnotes-module_item__LOUR3 > div")
    #若暫時無法開啟網頁 5秒後重新爬取一次
    if len(L_DT_data_d_ack) < 1:
        delay = random.uniform(2.0,6.0)
        sleep(delay)
        HP_DT_data = web_driver()
        HP_DT_data.get(data_url)
        HP_DT_data.execute_script("document.body.style.zoom='50%'")
        HP_DT_data.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
        L_DT_data_n = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
        L_DT_data_d_ack = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE")
        L_DT_specs = L_DT_soup.select("div.Container-module_root__luUPH.Container-module_container__jSUGk > div.Footnotes-module_item__LOUR3 > div")
        
    specs_data = []
    if len(L_DT_specs) > 0:
        L_DT_specs_list = L_DT_specs[-1].select("span")        
        if len(L_DT_specs_list) >0:
            for specs in L_DT_specs_list:
                specs_data.append(specs.text)
    HP_DT_data.quit()

    A = ["Model Name","Official Price","Web Link","Brand"]
    B = [ID_link[i],Money_link[i],data_url,"Hp"]    
    j = 0
    P = ""
    S = ""
    Wireless = ""
    WWAN = ""
    
    #依序抓取特徵名稱 ex: CPU
    for ack in L_DT_data_d_ack:
        L_DT_data_d = ack.select("div.Spec-module_valueWrapper__DTxWC > p.Typography-module_root__eQwd4.Typography-module_bodyM__XNddq.Spec-module_value__9FkNI.Typography-module_responsive__iddT7 > span")
        #抓取特徵內容
        N_D = L_DT_data_d[0].text
        if "[" in N_D and "]" in N_D:
            N_Dnum = N_D.split("[")[-1].split("]")[0]
            num = 0
            for num in range(len(N_Dnum.split(','))):
                sp_num = 0
                for sp_num in range(len(specs_data)):
                    if N_Dnum.split(',')[num] in specs_data[sp_num]:
                        N_D = N_D.replace(N_Dnum.split(',')[num], specs_data[sp_num].split("]")[-1]) 
        #篩選需求特徵 & 單位換算處理 & 資料切割
        if L_DT_data_n[j].text =="Processor" or L_DT_data_n[j].text =="Display" or L_DT_data_n[j].text =="Memory":
            A.append(L_DT_data_n[j].text)
            B.append(N_D)
        elif L_DT_data_n[j].text == "Operating system":
            A.append("Operating System")
            B.append(N_D) 
        elif L_DT_data_n[j].text == "Internal storage":
            A.append("Hard Drive")
            B.append(N_D)
        elif L_DT_data_n[j].text == "Audio Features":
            A.append("Audio and Speakers")
            B.append(N_D)
        elif L_DT_data_n[j].text == "Processor and graphics":
            A.append("Processor")
            A.append("Graphics Card")
            B.append(N_D.split("+")[0]) 
            B.append(N_D.split("+")[1])   
        elif L_DT_data_n[j].text == "Graphics":
            A.append("Graphics Card")
            B.append(N_D)
        elif L_DT_data_n[j].text == "Hard drive" or L_DT_data_n[j].text == "Storage":
            A.append("Hard Drive")
            B.append(N_D) 
        elif L_DT_data_n[j].text == "Wireless technology":
            Wireless = N_D  
        elif "WWAN" in L_DT_data_n[j].text:
            WWAN = N_D                      
        elif "Power" in L_DT_data_n[j].text:
            A.append("Power Supply")
            B.append(N_D)
        elif L_DT_data_n[j].text == "External I/O Ports":
            P = (N_D)
        elif L_DT_data_n[j].text == "Expansion slots":
            S = (N_D)            
        elif L_DT_data_n[j].text =="Dimensions (W X D X H)":
            Dim = (N_D).split("x")
            W = float(Dim[0].strip())*25.4
            D = float(Dim[1].strip())*25.4
            H = float(Dim[2].split("in")[0].strip())*25.4
            A.append("Width")
            A.append("Depth")
            A.append("Height")
            B.append(round(W,2))
            B.append(round(D,2))
            B.append(round(H,2))
        elif L_DT_data_n[j].text =="Weight":
            A.append("Weight")
            Weight_kg = float(N_D.split("lb")[0].strip())*0.4536            
            B.append(round(Weight_kg,2))
        j+=1
        
    #針對port與slot複合合併處理 
    PS = P+"\n"+S
    A.append("Ports & Slots")
    B.append(PS)
    
    #針對Wireless與WWAN複合合併處理 
    W_W = Wireless +"\n"+ WWAN
    A.append("Wireless")
    B.append(W_W)
    
    #資料合併
    A = pd.Series(A)
    if i == 0:
        C = pd.DataFrame(B,index = A)
    else:
        B = pd.DataFrame(B,index = A)
        C = C.merge(B,how = "outer",left_index=True, right_index=True)
        # C.drop_duplicates(inplace=True)    
C.to_excel("HP_DT.xlsx")