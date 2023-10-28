from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
import pandas as pd
import requests
import logging
import random

my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}
 
#設定特徵名稱
titles = ["Type","Brand","Model Name","Official Price","Ports & Slots","Display","Processor",'Dimensions and Weight',"Graphics Card","Storage","Memory","Operating System","Audio and Speakers",'Power',"Web Link"]
#設定網址
url = "https://www.dell.com/en-us/search/desktop?r=36679&p={}&ac=facetselect&t=Product"
DELL_DOCK_data = requests.get(url, headers=my_header)
sleep(2)

i = 1
ID_data = []
Money_data = []
link_data = []
Operating_System_data = []
Memory_data = []
Storage_data = []
Display_data = []
Processor_data = []
Graphics_data = []
D_W_dic = {}
# 確認連線狀況後讀取資料<連結>
if DELL_DOCK_data.status_code==200:
    L_NB_soup = BeautifulSoup(DELL_DOCK_data.text,"html.parser")
    sleep(2)
    #抓取商品數量<第一頁>
    
    all_number = L_NB_soup.select("p.pageinfo > label")    
    new_number = all_number[0].text.split("-")[-1].strip()
    tatle_number = all_number[1].text
    
    De_dock_data = L_NB_soup.select("article")        
    data = 0
    reset_OS,reset_Memory,reset_Storage,reset_Display,reset_Processor,reset_Graphics = "No Data","No Data","No Data","No Data","No Data","No Data"
    for data in De_dock_data:          
        D_deta_url = data.select("h3.ps-title > a")
        De_dock_money = data.select("div.ps-dell-price.ps-simplified")
        De_dock_title = data.select("span.ps-iconography-specs-title")
        De_dock_Data = data.select("span.ps-iconography-specs-label")
        if len(D_deta_url) > 0:
            #儲存連結
            ID_data.append(D_deta_url[0].text)
            Money_data.append(De_dock_money[0].text.split("$")[-1].strip())
            link_data.append("https:{}".format(D_deta_url[0]["href"]))
            Data = 0
            for Data in range(len(De_dock_title)):
                if "OS" in De_dock_title[Data].text:
                    reset_OS = De_dock_Data[Data].text.strip()
                elif "Memory" in De_dock_title[Data].text:
                    reset_Memory = De_dock_Data[Data].text.strip()
                elif "Storage" in De_dock_title[Data].text:
                    reset_Storage = De_dock_Data[Data].text.strip()
                elif "Display" in De_dock_title[Data].text:
                    reset_Display = De_dock_Data[Data].text.strip()                   
                elif "Processor" in De_dock_title[Data].text:
                    reset_Processor = De_dock_Data[Data].text.strip()                    
                elif "Graphics" in De_dock_title[Data].text:
                    reset_Graphic = De_dock_Data[Data].text.strip()
            Graphics_data.append(reset_Graphic)
            Processor_data.append(reset_Processor) 
            Operating_System_data.append(reset_OS)
            Memory_data.append(reset_Memory)
            Storage_data.append(reset_Storage)
            Display_data.append(reset_Display)                 
'''                 
#直到抓到的數量為0
i=i+1
while new_number != tatle_number:
    delay = random.uniform(1.0, 5.0)
    sleep(delay)
    url = "https://www.dell.com/en-us/search/desktop?r=36679&p={}&ac=facetselect&t=Product".format(i)
    i=i+1
    DELL_DOCK_data = requests.get(url, headers=my_header)
    sleep(2)
    if DELL_DOCK_data.status_code==200:
        L_NB_soup = BeautifulSoup(DELL_DOCK_data.text,"html.parser")
        sleep(2)
        all_number = L_NB_soup.select("p.pageinfo > label")    
        new_number = all_number[0].text.split("-")[-1].strip()
        #抓取特徵名稱
        De_dock_data = L_NB_soup.select("article")       
        data = 0
        reset_OS,reset_Memory,reset_Storage,reset_Display,reset_Processor,reset_Graphics = "No Data","No Data","No Data","No Data","No Data","No Data"
        for data in De_dock_data:          
            D_deta_url = data.select("h3.ps-title > a")
            De_dock_money = data.select("div.ps-dell-price.ps-simplified")
            De_dock_title = data.select("span.ps-iconography-specs-title")
            De_dock_Data = data.select("span.ps-iconography-specs-label")
            if len(D_deta_url) > 0:
                #儲存連結
                ID_data.append(D_deta_url[0].text)
                Money_data.append(De_dock_money[0].text.split("$")[-1].strip())
                link_data.append("https:{}".format(D_deta_url[0]["href"]))
                Data = 0
                for Data in range(len(De_dock_title)):
                    if "OS" in De_dock_title[Data].text:
                        reset_OS = De_dock_Data[Data].text.strip()
                    elif "Memory" in De_dock_title[Data].text:
                        reset_Memory = De_dock_Data[Data].text.strip()
                    elif "Storage" in De_dock_title[Data].text:
                        reset_Storage = De_dock_Data[Data].text.strip()
                    elif "Display" in De_dock_title[Data].text:
                        reset_Display = De_dock_Data[Data].text.strip()                   
                    elif "Processor" in De_dock_title[Data].text:
                        reset_Processor = De_dock_Data[Data].text.strip()                    
                    elif "Graphics" in De_dock_title[Data].text:
                        reset_Graphic = De_dock_Data[Data].text.strip()
                Graphics_data.append(reset_Graphic)
                Processor_data.append(reset_Processor) 
                Operating_System_data.append(reset_OS)
                Memory_data.append(reset_Memory)
                Storage_data.append(reset_Storage)
                Display_data.append(reset_Display)
'''                
j=0
#網頁爬取資料
for j in range(len(link_data)):
    print("Dell_DT {}".format(j))
    delay = random.uniform(0.5, 5.0)
    sleep(delay)
    url_dell = link_data[j] + "#techspecs_section"
    Type, Brand, Model_Name, Official_Price, Dimensions_Weight, Ports,Power_Supply, Keyboard, PalmRest = "","Dell",ID_data[j],Money_data[j],"","","","",""
    Slots, Ports_Slots, Camera, Display, Wireless, NFC, Primary_Battery, Processor, Graphics_Card = "","","",Display_data[j],"","","",Processor_data[j],Graphics_data[j]
    Storage, Memory, Operating_System, Audio_Speakers, User_guide, Web_Link = Storage_data[j],Memory_data[j],Operating_System_data[j],"","",link_data[j]
    FPR_model,FPR,Display_cleck = "","",""
    
    if D_W_dic.get(Model_Name) != None:
        Dimensions_Weight = D_W_dic.get(Model_Name)
    
    #開啟搜尋頁面
    option = webdriver.ChromeOptions()
    option.add_argument("headless")
    dell_dock = webdriver.Chrome(options=option)
    dell_dock.get(url_dell)
    
    dell_dock.execute_script("document.body.style.zoom='50%'")
    sleep(2)
    dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.5);")
    sleep(2)
    dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.25);")
    sleep(2)
    dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.65);")
    sleep(2)
    dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.4);")
    sleep(2)
    soup = BeautifulSoup(dell_dock.page_source,"html.parser")
    dell_dock.quit()
    #開始爬資料
    one_data = soup.select("ul.cf-hero-bts-list > li")
    for one in one_data:
        two_data = one.select("p")
        if "W " in two_data[0].text and "Windows" not in two_data[0].text:
            Power_Supply = two_data[0].text.strip()
        if "Ports" in two_data[0].text or "Slots" in two_data[0].text:
            two_data = one.select(" p > a")
            Ports_Slots = two_data[0]["data-description"]
            Ports_Slots = Ports_Slots.replace("<br>","\n")
            Ports_Slots = Ports_Slots.replace("<ul>","")
            Ports_Slots = Ports_Slots.replace("</ul>","")
            Ports_Slots = Ports_Slots.replace("<li>"," ")
            Ports_Slots = Ports_Slots.replace("</li>","\n")
        
    No_select_data = soup.select("li.mb-2")
    for no_data in No_select_data:
        No_select_title = no_data.select("div")
        No_select_Data = no_data.select("p")
        if "Ports" in No_select_title[0].text or "Slots" in  No_select_title[0].text:
            Ports_Slots = Ports_Slots + "\n" + No_select_Data[0].text
        elif ("Dimensions" in No_select_title[0].text or "Weight" in No_select_title[0].text) and Model_Name not in D_W_dic:
            Dimensions_Weight = No_select_Data[0].text
        elif "PalmRest" in No_select_title[0].text:
            PalmRest = No_select_Data[0].text
        elif "Camera" in No_select_title[0].text:
            Camera = No_select_Data[0].text
        elif "Audio Speakers" in No_select_title[0].text:
            Audio_Speakers = No_select_Data[0].text
        elif "Power" in No_select_title[0].text:
            Power_Supply = No_select_Data[0].text
        elif "Display" in No_select_title[0].text:
            Display_cleck= No_select_Data[0].text
            
    #針對無選擇規格<2st>
    Href = soup.select("div.ux-list-item-wrapper")
    for href in Href:
        Dell_tatle = href.select("div.ux-module-title-wrap > h2")
        dell_date = href.select("div.ux-readonly-title") 
        if len(Dell_tatle) > 0:
            if len(Power_Supply) < 5 and Dell_tatle[0].text.lower() == "power":
                Power_Supply = dell_date[0].text.lower()
                    
    #針對有選擇規格 <1st>              
    Href = soup.select("div.ux-cell-wrapper")
    for href in Href:
        Dell_tatle = href.select("div.ux-module-title-wrap > h2")
        if len(Dell_tatle) > 0:
            dell_date = href.select("div.ux-cell-options > div > div > div.ux-cell-title")
            dell_select = href.select("div.ux-cell-options > div > div > div.ux-cell-delta-price")
            number_data = 0
            for number_data in range(len(dell_date)):      
                if len(Power_Supply) < 10 and Dell_tatle[0].text.lower() == "power":
                    number_data = 0
                    for number_data in range(len(dell_date)):                
                        if "Select" in dell_select[number_data].text:
                            Power_Supply = dell_date[number_data].text.lower()                                        
                    
    #特殊爬取處理 滑動讀取資料迴圈
    if len(Ports_Slots) < 20 or len(Dimensions_Weight) < 20:
        Href = soup.select("div.pd-feature-wrap")
        P_S_data = ""
        
        number = 0    
        while number < 7:
            option = webdriver.ChromeOptions()                
            option.add_argument("headless")
            dell_dock = webdriver.Chrome(options=option)
            dell_dock.get(link_data[j] + '#ratings_section')

            sleep(2)
            # 獲取當前滾動位置
            scroll_height = dell_dock.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")                
            scroll_position = dell_dock.execute_script("return window.scrollY;")                
            # 計算滾動位置百分比 %數
            scroll_percentage = round((scroll_position / scroll_height),2)      
            dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*{});".format(scroll_percentage - number*0.05))
            sleep(2)
            soup = BeautifulSoup(dell_dock.page_source,"html.parser")
            Href = soup.select("div.pd-feature-wrap")
            dell_dock.quit()
            href_number = 0
            if len(Href) > 1:
                for href in Href:            
                    dell_tatle = href.select("h2")
                    href_number = href_number + 1
                    dell_tatle_2 = href.select("div > div > h2")
                    if len(dell_tatle) > 0:
                        if "Dimensions" in dell_tatle[0].text and len(Dimensions_Weight) < 20:
                            Dell_Con = ""
                            
                            con_num = 0    
                            dell_content = href.select("div.pd-item-desc")
                            for con_num in range (len(dell_content)):
                                Dell_Con = Dell_Con + dell_content[con_num].text.strip() +"\n"
                                data_NB = Dell_Con
                                
                            Dimensions_Weight = data_NB
   
                        if ("Ports" in dell_tatle[0].text or "Slots" in dell_tatle[0].text or "Connectivity" in dell_tatle[0].text) and len(P_S_data) < 20:
                            dell_date_PS = href.select("div.feature-col > div.pd-item-desc")
                            if len(dell_date_PS) > 0:
                                date = 0
                                for date in range(len(dell_date_PS)):
                                    P_S_data = P_S_data + "\n" + dell_date_PS[date].text
                                    P_S_data = P_S_data.strip()

                            dell_date_PS = Href[href_number].select("div > div.pd-item-desc")
                            if len(dell_date_PS) > 0:
                                date = 0
                                for date in range(len(dell_date_PS)):
                                    P_S_data = P_S_data + "\n" + dell_date_PS[date].text
                                    P_S_data = P_S_data.strip()                                
                            Ports_Slots = P_S_data  
                            
                    if len(dell_tatle_2) >0:
                        if "Dimensions" in dell_tatle_2[0].text and len(Dimensions_Weight) < 20:
                            Dell_Con = ""                                
                            con_num = 0    
                            dell_content = href.select("div.pd-item-desc")
                            for con_num in range (len(dell_content)):
                                Dell_Con = Dell_Con + dell_content[con_num].text.strip() +"\n"
                                data_NB = Dell_Con                                
                            Dimensions_Weight = data_NB
                            
            if len(Ports_Slots) > 20 and len(Dimensions_Weight) > 20:
                break
            number = number +1
    
    if len(Dimensions_Weight) < 20:
        Href = soup.select("ul.specs.list-unstyled")            
        while number < 9:
            option = webdriver.ChromeOptions()
            option.add_argument("headless")
            dell_dock = webdriver.Chrome(options=option)
            dell_dock.get(url_dell) 

            dell_dock.execute_script("document.body.style.zoom='50%'")
            dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*{});".format(0.4 + number*0.05))
            sleep(2)
            soup = BeautifulSoup(dell_dock.page_source,"html.parser")
            Href = soup.select("ul.specs.list-unstyled")
            dell_dock.quit()
            if len(Href) > 1:
                for href in Href:            
                    dell_tatle = href.select("li")
                    if len(dell_tatle) > 0:
                        for tatle in dell_tatle:
                            if "Dimensions" in tatle.select("div")[0].text.lower and len(Dimensions_Weight < 20):
                                data = tatle.select("p")
                                data_NB = data                              
                                Dimensions_Weight = data_NB
            if len(Dimensions_Weight) > 20:
                break
            number = number +1
    if Model_Name not in D_W_dic and len(Dimensions_Weight) > 20:
        D_W_dic[Model_Name] = Dimensions_Weight.strip()
        
    B = [Type, Brand, Model_Name, Official_Price, Ports_Slots.strip(), Display, Processor, Dimensions_Weight.strip(), Graphics_Card, Storage, Memory, Operating_System, Audio_Speakers,Power_Supply, Web_Link]
    A = pd.Series(titles)
    if j == 0:
        C = pd.DataFrame(B,index = A)
    else:
        B = pd.DataFrame(B,index = A)
        C = C.merge(B,how = "outer",left_index=True, right_index=True)

C = C.T
C.reset_index(drop = True, inplace = True)
re_rext = 0    
for re_rext in range(j+1):
    if len(C["Dimensions and Weight"][re_rext]) < 20 and D_W_dic.get(C["Model Name"][re_rext]) != None:
        C["Dimensions and Weight"][re_rext] = D_W_dic.get(C["Model Name"][re_rext])
        
C = C.T
C.to_excel("DELL_DT.xlsx")  
