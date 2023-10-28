from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import logging
import random

ID_link = []
Money_link = []
href_line = []

# 定義要爬取的網址
url = "https://www.hp.com/us-en/shop/sitesearch?keyword=Laptops"
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}   
#開啟搜尋頁面
HP_NB = webdriver.Chrome()
HP_NB.get(url)   
HP_NB.execute_script("document.body.style.zoom='50%'")
sleep(2)
'''
#頁面向下滾動 & 資料載入

while True:
    try:
        #向下滾動
        HP_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        HP_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        #按按鈕
        button = HP_NB.find_elements(By.CSS_SELECTOR,'button.hawksearch-load-more')
        HP_NB.execute_script("arguments[0].click();", button[0])
        sleep(2)
        HP_NB.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(3)
    except:
        break
'''
sleep(5)
#抓連結
soup = BeautifulSoup(HP_NB.page_source,'html.parser')
Href = soup.select("div.ProductTile-module_wrapper__hS8-C.productTile")
HP_NB.quit()

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
    href_line.append("https://www.hp.com"+ href_line_new[0]['href'])   

#分別進入各商品頁面
i = 0
for i in range(len(href_line)):
    print("HP_NB {}".format(i))
    delay = random.uniform(1.0, 5.0)
    sleep(delay)
    data_url = href_line[i]
    option = webdriver.ChromeOptions()
    HP_NB_data = webdriver.Chrome(options=option)   
    HP_NB_data.get(data_url + "#techSpecs")  
    HP_NB_data.execute_script("document.body.style.zoom='50%'")
    HP_NB_data.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(2)
    L_NB_soup = BeautifulSoup(HP_NB_data.page_source,'html.parser')
    HP_NB_data.quit()
    #抓取SPECS與OVERVIEW內文
    specs_data = []
    overview_data = []
    G4=""
    G5=""
    L_NB_specs = L_NB_soup.select("div.Container-module_root__luUPH.Container-module_container__jSUGk > div.Footnotes-module_item__LOUR3")
    for feature in range(len(L_NB_specs)):
        L_NB_feature = L_NB_specs[feature].select("div")
    
        if len(L_NB_feature) > 0:
            n = 0
            for n in range(len(L_NB_feature)):
                if L_NB_feature[n].text == "SPECS":
                    L_NB_specs_list = L_NB_specs[feature].select("span")
                    if len(L_NB_specs_list) >0:
                        for specs in L_NB_specs_list:
                            if specs != 0:
                                specs_data.append(specs.text)
                                if " 4G " in specs.text:
                                    G4 = "4G"
                                elif " 5G " in specs.text:
                                    G5 = "5G"
                    else:
                        L_NB_specs_list = L_NB_specs[feature].select("span.point")
                elif L_NB_feature[n].text == "OVERVIEW":
                    L_NB_overview_list = L_NB_specs[feature].select("span")        
                    if len(L_NB_overview_list) >0:
                        for overview in L_NB_overview_list:
                            if " 4G " in overview.text:
                                G4 = "4G"
                            elif " 5G " in overview.text:
                                G5 = "5G"
    if G4 != "" and G5 != "":
        WAN = "4G/5G"
    elif G4 != "" and G5 == "":
        WAN = "4G"
    elif G4 == "" and G5 != "":
        WAN = "5G"
    else:
        WAN = "No"

    L_NB_data_n = L_NB_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
    L_NB_data_d_ack = L_NB_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE")
    
    #針對不同網頁寫法抓取
    while len(L_NB_data_n) < 1:
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        option = webdriver.ChromeOptions()
        HP_NB_data = webdriver.Chrome(options=option)    
        HP_NB_data.get(data_url + "#techSpecs")

        HP_NB_data.execute_script("document.body.style.zoom='50%'")
        sleep(2)
        L_NB_soup = BeautifulSoup(HP_NB_data.page_source,'html.parser')               
        L_NB_data_n = L_NB_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
        L_NB_data_d_ack = L_NB_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE")
        HP_NB_data.quit()
 
    A = ["Model Name","Official Price","Web Link","Brand"]
    B = [ID_link[i],Money_link[i],data_url,"Hp"]
    j=0
    P=""
    S=""
    Wireless = "No"
    WWAN = ""
    FPR = "No"
    NFC = "No"
    FPR_model = "No"
    #依序抓取特徵名稱 ex: CPU
    for ack in L_NB_data_d_ack:
        L_NB_data_d = ack.select("div.Spec-module_valueWrapper__DTxWC > p.Typography-module_root__eQwd4.Typography-module_bodyM__XNddq.Spec-module_value__9FkNI.Typography-module_responsive__iddT7 > span")
        
        #抓取特徵內容
        N_D = L_NB_data_d[0].text
            
        if "[" in N_D and "]" in N_D:                
            N_Dnum = N_D.split("[")
            new_ND = ""
            number_ND = ""
            number_ND_data = ""
            for ND_number in range(len(N_Dnum)):
                new_ND = new_ND + N_Dnum[ND_number].split("]")[-1]
                number_ND = number_ND + N_Dnum[ND_number].split("]")[0]
            num = 0
            for num in range(len(number_ND.split(','))):
                sp_num = 0
                for sp_num in range(len(specs_data)):
                    if number_ND.split(',')[num] in specs_data[sp_num].split(']')[0]:
                        number_ND_data = number_ND_data + "\n" + specs_data[sp_num]
            N_D = new_ND + number_ND_data
                        
        
        #篩選需求特徵 & 單位換算處理 & 資料切割
        if L_NB_data_n[j].text == "Display" or L_NB_data_n[j].text == "Memory" or L_NB_data_n[j].text == "Processor":
            A.append(L_NB_data_n[j].text)
            B.append(N_D)
        elif "Finger" in L_NB_data_n[j].text:
            if "No" not in N_D:
                FPR = "Yes"
            if FPR == "Yes" or len(L_NB_data_d) > 1:
                FPR_model = "Yes"
        elif "NFC" in L_NB_data_n[j].text:
            NFC = "Yes"                
        elif L_NB_data_n[j].text == "Base features":
            N_D = N_D.split("with")[-1]
            if len(N_D.split(","))>1 :
                new_ND = N_D.split(",")
                ascort = 0
                for ascort in range(len(new_ND)):
                    if "RAM" in new_ND[ascort]:
                        A.append("Memory")
                        B.append(new_ND[ascort].split("and")[-1])
                    if "Graphics" in new_ND[ascort]:
                        A.append("Graphics Card")
                        B.append(new_ND[ascort].split("and")[-1])
                    if ("Intel" in new_ND[ascort] or "AMD" in new_ND[ascort]) and "Graphics" not in new_ND[ascort]:
                        A.append("Processor")
                        B.append(new_ND[ascort].split("and")[-1])
                    if ("storage" in new_ND[ascort] or "eMMC" in new_ND[ascort]) and "Hard Drive" not in A:
                        A.append("Hard Drive")
                        B.append(new_ND[ascort].split("and")[-1])
            else:
                new_ND = N_D.split("and")
                if len(new_ND) < 2:
                    new_ND = N_D.split("+")
                ascort = 0
                for ascort in range(len(new_ND)):
                    if "RAM" in new_ND[ascort] or 'memory' in new_ND[ascort]:
                        A.append("Memory")
                        B.append(new_ND[ascort])
                    if "Graphics" in new_ND[ascort]:
                        A.append("Graphics Card")
                        B.append(new_ND[ascort])
                    if ("Intel" in new_ND[ascort] or "AMD" in new_ND[ascort]) and "Graphics" not in new_ND[ascort]:
                        A.append("Processor")
                        B.append(new_ND[ascort])
                    if "storage" in new_ND[ascort] and "Hard Drive" not in A:
                        A.append("Hard Drive")
                        B.append(new_ND[ascort])
        elif L_NB_data_n[j].text == "Operating system":
            A.append("Operating System")
            B.append(N_D) 
        elif L_NB_data_n[j].text == "Internal storage":
            A.append("Hard Drive")
            B.append(N_D)
        elif L_NB_data_n[j].text == "Processor and graphics":
            A.append("Processor")
            A.append("Graphics Card")
            B.append(L_NB_data_d[0].text.split("+")[0])
            B.append(N_D.split("+")[1])
        elif L_NB_data_n[j].text == "Processor, graphics, memory & hard disk":
            A.append("Processor")
            A.append("Graphics Card")
            A.append("Memory")            
            B.append(N_D.split("+")[0])
            B.append(N_D.split("+")[1])
            B.append(N_D.split("+")[2])
            if len(N_D.split("+"))>3:
                A.append("Hard Drive")
                B.append(L_NB_data_d[0].text.split("+")[3])
        elif L_NB_data_n[j].text == "Processor, graphics & memory":
            A.append("Processor")
            A.append("Graphics Card")
            A.append("Memory")
            B.append(N_D.split("+")[0])
            B.append(N_D.split("+")[1])
            B.append(N_D.split("+")[2])
        elif L_NB_data_n[j].text == "Graphics":
            if "Graphics Card" not in A:
                A.append("Graphics Card")
                B.append(N_D.split(":")[-1])
        elif "Power" in L_NB_data_n[j].text:
            A.append("Power Supply")
            B.append(N_D)
        elif L_NB_data_n[j].text == "Storage" or L_NB_data_n[j].text == "Hard drive" or  L_NB_data_n[j].text == "Internal drive":
            A.append("Hard Drive")
            B.append(N_D)
        elif L_NB_data_n[j].text == "Primary battery" or L_NB_data_n[j].text == "Battery":
            A.append("Primary Battery")
            B.append(N_D)
        elif L_NB_data_n[j].text == "Webcam" or L_NB_data_n[j].text == "Camera":
            A.append("Camera")
            B.append(N_D)
        elif L_NB_data_n[j].text == "Expansion slots":
            S = (N_D)            
        elif "Ports" in L_NB_data_n[j].text or L_NB_data_n[j].text == "External I/O Ports":
            P = (N_D)
        elif L_NB_data_n[j].text == "Weight":
            A.append("Weight")
            Weight_kg_A = N_D.split("lb")[0].split("at")[-1].strip()
            if "<" in Weight_kg_A:
                Weight_kg_A = Weight_kg_A.split("<")[-1]
                Weight_kg = float(Weight_kg_A)*0.4536            
                B.append("< {}".format(round(Weight_kg,2)))
            else:
                Weight_kg = float(Weight_kg_A)*0.4536            
                B.append(round(Weight_kg,2))
        elif L_NB_data_n[j].text == "Dimensions (W X D X H)":
            Dim = (N_D).split("x")
            W = float(Dim[0].split("at")[-1].strip())*25.4
            D = float(Dim[1].strip())*25.4
            H_real = Dim[2].split("in")[0].strip().split("-")
            if len(H_real)>1:
                h,H = 0,0
                for h in range(len(H_real)):
                    if H < float(H_real[h].split("(")[0].strip()):
                        H = float(H_real[h].split("(")[0].strip())
                    else:
                        pass
                H = H*25.4
            else:
                H = float(Dim[2].split("in")[0].strip())*25.4
            A.append("Width")
            A.append("Depth")
            A.append("Height")
            B.append(round(W,2))
            B.append(round(D,2))
            B.append(round(H,2))
        elif L_NB_data_n[j].text ==  "Audio Features":
            A.append("Audio and Speakers")
            B.append(N_D)
        j+=1
        
    A.append("FPR")
    B.append(FPR)
    A.append("FPR_model")
    B.append(FPR_model)
    A.append("NFC")
    B.append(NFC)
    
    #針對port與slot複合合併處理
    PS = P+"\n"+S
    A.append("Ports & Slots")
    B.append(PS)
    
    A.append("Wireless")
    B.append(WAN)
    
    #資料合併
    A = pd.Series(A)
    if i == 0:
        C = pd.DataFrame(B,index = A)
    else:
        B = pd.DataFrame(B,index = A)
        C = C.merge(B,how = "outer",left_index=True, right_index=True)

C.to_excel("HP_NB.xlsx")

