from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
import pandas as pd
import requests
import logging
import random
from fake_useragent import UserAgent

# 配置日志
logging.basicConfig(filename='bug.log',
                    filemode='w',
                    datefmt='%a %d %b %Y %H:%M:%S',
                    format='%(asctime)s %(filename)s %(levelname)s:%(message)s',
                    level=logging.INFO)

#讀取下載的IP代理位置檔案
proxy = pd.read_csv('proxyscrape_premium_http_proxies.txt', delimiter='\t', encoding='utf-8',header=None)
proxy_data = proxy.values.tolist()

try:    
    my_header = {'user-agent':UserAgent().random}
    #設定特徵名稱
    titles = ["Type","Brand","Model Name","Official Price","Ports & Slots","Camera","Display","Primary Battery","Processor","Graphics Card","Storage","Memory","Operating System","Audio and Speakers","Dimensions and Weight","Wireless","NFC","FPR","FPR_model",'Power',"Web Link"]
    #設定網址
    url = "https://www.dell.com/en-us/search/laptop?r=36679&p={}&ac=facetselect&t=Product"
    DELL_DOCK_data = requests.get(url, headers=my_header)
    sleep(2)
    
    ip_number = 0
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
        all_number = L_NB_soup.select("p.pageinfo > label")    
        new_number = all_number[0].text.split("-")[-1].strip()
        tatle_number = all_number[1].text
        #抓取商品數量<第一頁>
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
      
    #直到抓到的數量為0
    i=i+1
    while new_number != tatle_number:
        delay = random.uniform(1.0, 5.0)
        sleep(delay)
        url = "https://www.dell.com/en-us/search/laptop?r=36679&p={}&ac=facetselect&t=Product".format(i)
        i=i+1
        DELL_DOCK_data = requests.get(url, headers=my_header)
        sleep(2)
        if DELL_DOCK_data.status_code==200:
            L_NB_soup = BeautifulSoup(DELL_DOCK_data.text,"html.parser")
            all_number = L_NB_soup.select("p.pageinfo > label")    
            new_number = all_number[0].text.split("-")[-1].strip()
            sleep(3)
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
    new_ip = random.choice(proxy_data)[0]
    
    j=0
    for j in range(len(link_data)):
        print("Dell_NB {}".format(j))
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        url_dell = link_data[j] + "#techspecs_section"
        Type, Brand, Model_Name, Official_Price, Dimensions_Weight, Ports,Power_Supply, Keyboard, PalmRest = "","Dell",ID_data[j],Money_data[j],"","","","",""
        Slots, Ports_Slots, Camera, Display, Wireless, NFC, Primary_Battery, Processor, Graphics_Card = "","","",Display_data[j],"","","",Processor_data[j],Graphics_data[j]
        Storage, Memory, Operating_System, Audio_Speakers, User_guide, Web_Link = Storage_data[j],Memory_data[j],Operating_System_data[j],"","",link_data[j]
        FPR_model,FPR,Display_cleck = "","",""
        
        if D_W_dic.get(Model_Name) != None:
            Dimensions_Weight = D_W_dic.get(Model_Name)
            sleep(2)
        
        #開啟搜尋頁面
        option = webdriver.ChromeOptions()
        #使用代理
        if ip_number == 50:
            new_ip = random.choice(proxy_data)[0]
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            ip_number = 0
        else:
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            
        option.add_argument("headless")
        dell_dock = webdriver.Chrome(options=option)
        dell_dock.get(url_dell)
        sleep(2)
        while "Access Denied" in dell_dock.page_source:
            dell_dock.quit()
            sleep(2)
            new_ip = random.choice(proxy_data)[0]
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            option.add_argument("headless")
            dell_dock = webdriver.Chrome(options=option)
            dell_dock.get(url_dell)
            sleep(2)
            ip_number = 0
        dell_dock.execute_script("document.body.style.zoom='50%'")
        sleep(2)
        dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.2);")
        sleep(2)
        dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.5);")
        sleep(2)
        dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.8);")
        sleep(2)
        soup = BeautifulSoup(dell_dock.page_source,"html.parser")
        dell_dock.quit()
        
        #開始爬資料
        one_data = soup.select("ul.cf-hero-bts-list > li")
        for one in one_data:
            two_data = one.select("p")
            if "Power Supply" in two_data[0].text:
                Power_Supply = two_data[0].text.strip()
            if "Ports" in two_data[0].text or "Slots" in two_data[0].text or "PORTS" in two_data[0].text:
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
            if "Ports" in No_select_title[0].text or "Slots" in No_select_title[0].text or "PORTS" in No_select_title[0].text:
                Ports_Slots = Ports_Slots + "\n" + No_select_Data[0].text
            elif ("Dimensions" in No_select_title[0].text or "Weight" in No_select_title[0].text) and Model_Name not in D_W_dic:
                Dimensions_Weight = No_select_Data[0].text
            elif "Keyboard" in No_select_title[0].text:
                Keyboard = No_select_Data[0].text
            elif "PalmRest" in No_select_title[0].text:
                PalmRest = No_select_Data[0].text
            elif "Camera" in No_select_title[0].text:
                Camera = No_select_Data[0].text
            elif "Audio Speakers" in No_select_title[0].text:
                Audio_Speakers = No_select_Data[0].text
            elif "NFC" in No_select_title[0].text:
                NFC = No_select_Data[0].text
            elif "Primary Battery" in No_select_title[0].text:
                Primary_Battery = No_select_Data[0].text
            elif "Power" in No_select_title[0].text and len(Power_Supply) < 10:
                Power_Supply = No_select_Data[0].text
            elif "User guide" in No_select_title[0].text:
                User_guide = No_select_Data[0].text
            elif "Display" in No_select_title[0].text:
                Display_cleck= No_select_Data[0].text        
    
        #NFC FPR Wireless處理
        KP = Keyboard + PalmRest
        KP_model = Keyboard + PalmRest
        FPR = ""
        NFC = ""
        #針對無選擇規格<1st>
        Href = soup.select("li.mb-2")
        for href in Href:
            dell_tatle = href.select("div")
            dell_date = href.select("p")
            if len(dell_tatle) > 0 and len(dell_date) > 0:
                if dell_tatle[0].text.lower() == "keyboard":
                    KP = KP +" "+ dell_date[0].text.lower()
                    KP_model = KP_model + " " + dell_date[0].text.lower()
                if dell_tatle[0].text.lower() == "palmrest":
                    KP = KP +" "+ dell_date[0].text.lower()
                    KP_model = KP_model + " " + dell_date[0].text.lower()
                    
        #針對無選擇規格<2st>
        Href = soup.select("div.ux-list-item-wrapper")
        for href in Href:
            Dell_tatle = href.select("div.ux-module-title-wrap > h2")
            dell_date = href.select("div.ux-readonly-title") 
            if len(Dell_tatle) > 0:
                if Dell_tatle[0].text.lower() =="palmrest" or Dell_tatle[0].text.lower() =="keyboard":                
                    KP_model = KP_model + " " + dell_date[0].text.lower()
                    KP = KP +" "+ dell_date[0].text.lower()
                if len(Camera) < 5 and Dell_tatle[0].text.lower() =="camera":
                    Camera = dell_date[0].text.lower()
                if len(Power_Supply) < 10 and "power" in Dell_tatle[0].text.lower():
                    Power_Supply = dell_date[0].text.lower()
                if len(Primary_Battery) < 5 and Dell_tatle[0].text.lower() == "primary battery":
                    Primary_Battery = dell_date[0].text.lower()
                    
        #針對有選擇規格 <1st>              
        Href = soup.select("div.ux-cell-wrapper")
        for href in Href:
            Dell_tatle = href.select("div.ux-module-title-wrap > h2")
            if len(Dell_tatle) > 0:
                dell_date = href.select("div.ux-cell-options > div > div > div.ux-cell-title")
                dell_select = href.select("div.ux-cell-options > div > div > div.ux-cell-delta-price")
                number_data = 0
                for number_data in range(len(dell_date)):
                    if Dell_tatle[0].text.lower() =="palmrest" or Dell_tatle[0].text.lower() =="keyboard":
                        KP_model = KP_model + " " + dell_date[number_data].text.lower()
                        if "Select" in dell_select[number_data].text:
                            KP = KP +" "+ dell_date[number_data].text.lower()
                    if len(Camera) < 10 and Dell_tatle[0].text.lower() == "camera":
                        number_data = 0
                        for number_data in range(len(dell_date)):                
                            if "Select" in dell_select[number_data].text:
                                Camera = dell_date[number_data].text.lower()    
                    if len(Power_Supply) < 10 and "power" in Dell_tatle[0].text.lower():
                        number_data = 0
                        for number_data in range(len(dell_date)):                
                            if "Select" in dell_select[number_data].text:
                                Power_Supply = dell_date[number_data].text.lower()
                    if len(Primary_Battery) < 10 and Dell_tatle[0].text.lower() == "primary battery":
                        number_data = 0
                        for number_data in range(len(dell_date)):                
                            if "Select" in dell_select[number_data].text:
                                Primary_Battery = dell_date[number_data].text.lower()            
                                                                
        #針對有選擇規格 <2nd>         
        Href = soup.select("div.accordion > div")       
        for href in Href:
            dell_tatle = href.select("h2")
            dell_date = href.select("span")
            if len(dell_tatle) > 0 and len(dell_date) > 0:
                if dell_tatle[0].text.lower() == "keyboard":
                    KP = KP +" "+ dell_date[0].text.lower()
                    KP_model = KP_model + " " + dell_date[0].text.lower()
                if dell_tatle[0].text.lower() == "palmrest":
                    KP = KP +" "+ dell_date[0].text.lower()
                    KP_model = KP_model + " " + dell_date[0].text.lower()
                if len(Power_Supply) < 5 and "power" in dell_tatle[0].text.lower():
                    Power_Supply = dell_date[0].text
                if len(Primary_Battery) < 5 and "battery" in dell_tatle[0].text.lower():
                    Primary_Battery = dell_date[0].text
        
        if "fingerprint" in str(KP_model) or "fpr" in str(KP_model):
            FPR_model = "Yes"
        else:
            FPR_model = "No"
                
        if ("fingerprint" in str(KP) and "no fingerprint" not in str(KP)) or ("fpr" in str(KP) and "no fpr" not in str(KP)):
            FPR = "Yes"
        else:
            FPR = "No"
        
        if "nfc" in str(KP) and "no nfc" not in str(KP):
            NFC = "Yes"
        else:
            NFC = "No"
                
        WWAN_txt = ""
        #無選擇規格
        WWAN_data = soup.select("div.pd-item-desc > ul > li")
        if len(WWAN_data) < 1:
            WWAN_data = soup.select("div.pd-item-desc")
            for wwan in range(len(WWAN_data)):
                WWAN_txt = WWAN_txt + WWAN_data[wwan].text
        #有選擇規格
        WWAN_data = soup.select("div.ux-module-row-wrap > div")
        for wwan in WWAN_data:
            WWAN_data1 = wwan.select("h2")
            if WWAN_data1[0].text.lower() == "mobile broadband":
                WWAN_data2 = wwan.select("div.ux-cell-title")
                WWAN_select = wwan.select("div.ux-cell-delta-price")
                for wwan1 in range(len(WWAN_data2)):
                    if "select" in WWAN_select[wwan1].text.lower():
                        WWAN_txt = WWAN_txt + WWAN_data2[wwan1].text
                        
        if ("4G" in WWAN_txt and "5G" in WWAN_txt) or ("4G" in str(Display_cleck) and "5G" in str(Display_cleck)):
            Wireless = "4G/5G"
        elif ("4G" in WWAN_txt and "5G" not in WWAN_txt) or ("4G" in str(Display_cleck) and "5G" not in str(Display_cleck)):
            Wireless = "4G"
        elif ("4G" not in WWAN_txt and "5G" in WWAN_txt) or ("4G" not in str(Display_cleck) and "5G" in str(Display_cleck)):
            Wireless = "5G"
        else:
            Wireless = "No"    
        ip_number = ip_number+1    
        #特殊爬取處理 滑動讀取資料迴圈    
        if len(Ports_Slots) < 20 or len(Dimensions_Weight) < 20:
            Href = soup.select("div.pd-feature-wrap")
            P_S_data = ""
            number = 0    
            while number < 11:
                option = webdriver.ChromeOptions()
                
                #使用代理
                if ip_number == 50:
                    new_ip = random.choice(proxy_data)[0]
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                    ip_number = 0
                else:
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                    
                option.add_argument("headless")
                dell_dock = webdriver.Chrome(options=option)
                dell_dock.get(link_data[j] + '#features_section')
                sleep(2)
                
                while "Access Denied" in dell_dock.page_source:
                    dell_dock.quit()
                    sleep(2)
                    new_ip = random.choice(proxy_data)[0]
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                    option.add_argument("headless")
                    dell_dock = webdriver.Chrome(options=option)
                    dell_dock.get(link_data[j] + '#features_section')
                    sleep(2)
                    ip_number = 0                    
                dell_dock.execute_script("document.body.style.zoom='50%'")
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
                href_number = 0
                dell_dock.quit()
                ip_number = ip_number+1
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
                                
                            if ("Ports" in dell_tatle[0].text or "Slots" in dell_tatle[0].text or "PORTS" in dell_tatle[0].text) and len(P_S_data) < 20:
                                dell_date_PS = href.select("div.pd-item-desc")
                                if len(dell_date_PS) > 0:  
                                    date = 0
                                    for date in range(len(dell_date_PS)):
                                        P_S_data = P_S_data + "\n" + dell_date_PS[date].text
                                        P_S_data = P_S_data.strip()
                                if len(P_S_data) < 20:
                                    P_S_two = Href[href_number].select("div.pd-item-desc")
                                    if len(P_S_two) > 0:  
                                        date = 0
                                        for date in range(len(P_S_two)):
                                            P_S_data = P_S_data + "\n" + P_S_two[date].text
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
                #使用代理
                if ip_number == 50:
                    new_ip = random.choice(proxy_data)[0]
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                    ip_number = 0
                else:
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                option.add_argument("headless")
                dell_dock = webdriver.Chrome(options=option)
                dell_dock.get(url_dell)
                sleep(2)
                while "Access Denied" in dell_dock.page_source:
                    dell_dock.quit()
                    sleep(2)
                    new_ip = random.choice(proxy_data)[0]
                    random_proxy = new_ip
                    option.add_argument("--proxy-server=http://"+random_proxy)
                    option.add_argument("headless")
                    dell_dock = webdriver.Chrome(options=option)
                    dell_dock.get(link_data[j] + '#features_section')
                    sleep(2)
                    ip_number = 0               
                ip_number = ip_number+1
                sleep(2)
                # 獲取當前滾動位置
                scroll_height = dell_dock.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")                
                scroll_position = dell_dock.execute_script("return window.scrollY;")                
                # 計算滾動位置百分比 %數
                scroll_percentage = round((scroll_position / scroll_height),2)      
                dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*{});".format(scroll_percentage - number*0.05))

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
        
        if len(Ports_Slots) > 10 :
            Ports_Slots = Ports_Slots.strip()
            P_S_new = Ports_Slots.split("<span")
            if len(P_S_new) > 1:
                data_number = ""
                for Number_PS in range(len(P_S_new)):
                    if "span>" in P_S_new[Number_PS]:
                        data_number = data_number + P_S_new[Number_PS].split("span>")[-1]
                    else:
                        data_number = data_number + P_S_new[Number_PS]
                Ports_Slots = data_number
            Ports_Slots = Ports_Slots.replace("<p>"," ")
            Ports_Slots = Ports_Slots.replace("</p>","\n")
            Ports_Slots = Ports_Slots.replace("</span>","")
            Ports_Slots = Ports_Slots.replace("<span>","")
            Ports_Slots = Ports_Slots.replace("&amp;","")            
        B = [Type, Brand, Model_Name, Official_Price, Ports_Slots, Camera, Display, Primary_Battery, Processor, Graphics_Card, Storage, Memory, Operating_System, Audio_Speakers, Dimensions_Weight.strip(), Wireless, NFC, FPR, FPR_model,Power_Supply, Web_Link]
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
    # C.to_excel("DELL_NB.xlsx")
except Exception as bug:
    # 捕获并记录错误日志
    logging.error(f"An error occurred: {str(bug)}", exc_info=True)   
