import pandas as pd
import xlsxwriter
import logging
import random
from time import sleep
from selenium import webdriver
import requests
from bs4 import BeautifulSoup

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# 配置日志
logging.basicConfig(filename='bug.log',
                    filemode='w',
                    datefmt='%a %d %b %Y %H:%M:%S',
                    format='%(asctime)s %(filename)s %(levelname)s:%(message)s',
                    level=logging.INFO)

# try:
#laptop資料重整
df = pd.read_excel("DELL_NB.xlsx",index_col=0)

df = df.T
df.reset_index(drop = True, inplace = True)
df.rename(columns={'Power': 'Power Supply'}, inplace=True)
df.rename(columns={'Storage': 'Hard Drive'}, inplace=True)
service = Service(ChromeDriverManager().install())
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}

#將商品名稱與長寬深/重量抓出建立字典<方便後續補齊 減少需抓捕資料>
Dell_NB_data_DW = {}
Dell_NB_data_camera = {}
DNB = 0
for DNB in range(len(df["Brand"])):
    if df["Model Name"][DNB] not in Dell_NB_data_DW and len(str(df["Dimensions and Weight"][DNB])) > 25 :
        Dell_NB_data_DW[df["Model Name"][DNB]] = str(df["Dimensions and Weight"][DNB])        
    if df["Model Name"][DNB] not in Dell_NB_data_camera and len(str(df["Camera"][DNB])) > 10 :
        Dell_NB_data_camera[df["Model Name"][DNB]] = str(df["Camera"][DNB])  

DNB = 0
for DNB in range(len(df["Brand"])):
    if len(str(df["Camera"][DNB])) < 10 and df["Model Name"][DNB] in Dell_NB_data_camera:
        df["Camera"][DNB] = Dell_NB_data_camera.get(df["Model Name"][DNB])

Height = []
Width = []
Depth = []
Weight = []

#Dell_NB 資料檢驗/補充
DNB = 0
for DNB in range(len(df["Brand"])):        
    if len(str(df["Ports & Slots"][DNB])) < 20:
        print(DNB)
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        url_dell = df["Web Link"][DNB] + "#techspecs_section"
        option = webdriver.ChromeOptions()
        option.add_argument("headless")
        dell_dock = webdriver.Chrome(options=option)
        dell_dock.get(url_dell)
        sleep(2)
        dell_dock.execute_script("document.body.style.zoom='50%'")
        sleep(2)
        dell_dock.execute_script("window.scrollTo(0, document.body.scrollHeight*0.5);")
        sleep(2)
        soup = BeautifulSoup(dell_dock.page_source,"html.parser")
        dell_dock.quit()
        one_data = soup.select("ul.cf-hero-bts-list > li")
        Ports_Slots=""
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
                Ports_Slots = Ports_Slots.replace("<span>"," ")
                Ports_Slots = Ports_Slots.replace("</span>","")
        No_select_data = soup.select("li.mb-2")
        for no_data in No_select_data:
            No_select_title = no_data.select("div")
            No_select_Data = no_data.select("p")
            if "Ports" in No_select_title[0].text or "Slots" in No_select_title[0].text or "PORTS" in No_select_title[0].text:
                Ports_Slots = Ports_Slots + "\n" + No_select_Data[0].text
        df["Ports & Slots"][DNB] = Ports_Slots
#針對DELL重量體積資料處理 & 單位換算 & 資料切割        
S_W = df["Dimensions and Weight"]
i = 0
for i in range(len(S_W)):
    if str(Dell_NB_data_DW.get(df["Model Name"][i])) != "None":
        A00 = Dell_NB_data_DW[df["Model Name"][i]].lower()
        A0 = A00.split("weight")
        no_weight = 0
        for no_weight in range(len(A0)):
            height_data = " "
            width_data = " "
            depth_data = " "
            if "weight" not in A0[no_weight].lower() and "kg" not in A0[no_weight].lower() and ("mm" in A0[no_weight].lower() or "cm" in A0[no_weight].lower()):
                A1_text = A0[no_weight].split("width")
                A1_besa = A0[no_weight]                   
            if len(A1_text) <= 2:
                if "mm" in A1_besa:
                    if "h/w/d" in A1_besa:
                        A1_hwd = A1_besa.split("(")
                        hwd = 0
                        for hwd in range(len(A1_hwd)):
                            if "mm" in A1_hwd[hwd]:
                                A1_new_besa = A1_hwd[hwd]
                                A1 = A1_new_besa.split("x")
                                height_data = A1[0].split("height")[-1].split("3. ")[0].split("2. ")[0].split("(")[-1].split("-")[-1].split("–")[-1].split(":")[-1].split("mm")[0].strip()
                                width_data = A1[1].split("width")[-1].split("3. ")[0].split("2. ")[-1].split("(")[-1].split(":")[-1].split("mm")[0].strip()
                                depth_data = A1[2].split("depth")[-1].split("length")[-1].split("3. ")[-1].split("(")[-1].split(":")[-1].split("mm")[0].strip()
                    else:            
                        A1 = A1_besa.split("mm")            
                        mm_data = 0
                        for mm_data in range(len(A1)):
                            if "height" in A1[mm_data]:
                                height_data = A1[mm_data].split("height")[-1].split("3. ")[0].split("2. ")[0].split("(")[-1].split("-")[-1].split("–")[-1].split(":")[-1].strip()
                            elif "width" in A1[mm_data]:
                                width_data = A1[mm_data].split("width")[-1].split("3. ")[0].split("2. ")[-1].split("(")[-1].split(":")[-1].strip()
                            elif "depth" in A1[mm_data] or "length" in A1[mm_data]:
                                depth_data = A1[mm_data].split("depth")[-1].split("length")[-1].split("3. ")[-1].split("(")[-1].split(":")[-1].strip()
                elif "cm" in A1_besa:
                    A1 = A1_besa.split("cm")
                    cm_data = 0
                    for cm_data in range(len(A1)):
                        if "height" in A1[cm_data]:
                            height_data = float(A1[cm_data].split("height")[-1].split("3. ")[0].split("2. ")[0].split("(")[-1].split(")")[0].split("-")[-1].split(":")[-1].strip())*10
                        elif "width" in A1[cm_data]:
                            width_data = float(A1[cm_data].split("width")[-1].split("3. ")[0].split("2. ")[-1].split("(")[-1].split(":")[-1].strip())*10
                        elif "depth" in A1[cm_data] or "length" in A1[cm_data]:
                            depth_data = float(A1[cm_data].split("depth")[-1].split("length")[-1].split("3. ")[-1].split("(")[-1].split(":")[-1].strip())*10
            Weight_data = " "
            if len(A0) >1:
                if len(df["Model Name"][i].split("Special")) > 1:
                    if "kg" in A0[-2]:
                        A2 = A0[-2].split("kg")
                        Weight_data = A2[0].split("(")[-1].split(":")[-1].strip()
                    elif "g" in A0[-2]:
                        A2 = A0[-2].split("g")
                        Weight_data = float(A2[0].split("(")[-1].split(":")[-1].strip())/1000
                else:
                    if "kg" in A0[-1]:
                        A2 = A0[-1].split("kg")
                        Weight_data = A2[0].split("(")[-1].split(":")[-1].strip()
                    elif "g" in A0[-1]:
                        A2 = A0[-1].split("g")
                        Weight_data = float(A2[0].split("(")[-1].split(":")[-1].strip())/1000                
        else:
            pass
        Height.append(height_data)
        Width.append(width_data)
        Depth.append(depth_data)
        Weight.append(Weight_data)
    else:
        Height.append(" ")        
        Width.append(" ")
        Depth.append(" ")
        Weight.append(" ")

#重新填充資料欄位與去除欄位
df["Height"] = Height
df["Width"] = Width
df["Depth"] = Depth
df["Weight"] = Weight

df = df.sort_values(["Model Name"],ascending=True)
df = df.T

#載入其他公司資料進行合併

df_1 = pd.read_excel("HP_NB.xlsx",index_col="Unnamed: 0")
df_2 = pd.read_excel("Lenovo_NB.xlsx",index_col="Unnamed: 0")

df_1 = df_1.T
df_2 = df_2.T
df_1= df_1.reset_index(drop=True)
df_2= df_2.reset_index(drop=True)

HP_NB_data_camera = {}
HP_NB_data_OS = {}
HP_NB_data_Processor = {}
HP_NB_data_GC = {}

#HP_NB 現有資料剪貼
DNB = 0
for DNB in range(len(df_1["Brand"])):
    if len(str(df_1["Camera"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_camera:
        df_1["Camera"][DNB] = HP_NB_data_camera.get(df_1["Model Name"][DNB])
    if len(str(df_1["Operating System"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_OS:
        df_1["Operating System"][DNB] = HP_NB_data_OS.get(df_1["Model Name"][DNB])        
    if len(str(df_1["Processor"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_Processor:
        df_1["Processor"][DNB] = HP_NB_data_Processor.get(df_1["Model Name"][DNB])        
    if len(str(df_1["Graphics Card"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_GC:
        df_1["Graphics Card"][DNB] = HP_NB_data_GC.get(df_1["Model Name"][DNB])
    if len(str(df_1["Audio and Speakers"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_GC:
        df_1["Audio and Speakers"][DNB] = HP_NB_data_GC.get(df_1["Model Name"][DNB])
        
DNB = 0
for DNB in range(len(df_1["Brand"])):
    if len(str(df_1["Camera"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_camera:
        df_1["Camera"][DNB] = HP_NB_data_camera.get(df_1["Model Name"][DNB])
    if len(str(df_1["Operating System"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_OS:
        df_1["Operating System"][DNB] = HP_NB_data_OS.get(df_1["Model Name"][DNB])
    if len(str(df_1["Processor"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_Processor:
        df_1["Processor"][DNB] = HP_NB_data_Processor.get(df_1["Model Name"][DNB])    
    if len(str(df_1["Graphics Card"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_GC:
        df_1["Graphics Card"][DNB] = HP_NB_data_GC.get(df_1["Model Name"][DNB])
    if len(str(df_1["Audio and Speakers"][DNB])) < 10 and df_1["Model Name"][DNB] in HP_NB_data_GC:
        df_1["Audio and Speakers"][DNB] = HP_NB_data_GC.get(df_1["Model Name"][DNB])

Lenovo_NB_data_camera = {}
Lenovo_NB_data_PB = {}
Lenovo_NB_data_PS = {}
     
#Lenovo_NB 現有資料剪貼
DNB = 0
for DNB in range(len(df_2["Brand"])):
    if len(str(df_2["Camera"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_camera:
        df_2["Camera"][DNB] = Lenovo_NB_data_camera.get(df_2["Model Name"][DNB])
    if len(str(df_2["Primary Battery"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_PB:
        df_2["Primary Battery"][DNB] = Lenovo_NB_data_PB.get(df_2["Model Name"][DNB])        
    if len(str(df_2["Ports & Slots"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_PS:
        df_2["Ports & Slots"][DNB] = Lenovo_NB_data_PS.get(df_2["Model Name"][DNB])        

DNB = 0
for DNB in range(len(df_2["Brand"])):
    if len(str(df_2["Audio and Speakers"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_camera:
        df_2["Audio and Speakers"][DNB] = Lenovo_NB_data_camera.get(df_2["Model Name"][DNB])
    if len(str(df_2["Primary Battery"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_PB:
        df_2["Primary Battery"][DNB] = Lenovo_NB_data_PB.get(df_2["Model Name"][DNB])
    if len(str(df_2["Ports & Slots"][DNB])) < 10 and df_2["Model Name"][DNB] in Lenovo_NB_data_PS:
        df_2["Ports & Slots"][DNB] = Lenovo_NB_data_PS.get(df_2["Model Name"][DNB])

df_2= df_2.reset_index(drop=True)
re_load = 0
# 檢查Lenovo_NB的缺值(HWD/Weight)並補齊
for re_load in range(len(df_2)):
    if str(df_2['Depth'][re_load]) == "nan" :
        print(re_load)
        delay = random.uniform(1.0, 5.0)
        sleep(delay)
        data_url = df_2['Web Link'][re_load] 
        Lenovo_NB_data = requests.get(data_url + "#features", headers=my_header)
        if Lenovo_NB_data.status_code==200:
            L_NB_soup = BeautifulSoup(Lenovo_NB_data.text,'html.parser')
            #商品詳細網址            
            NB_deta_url = L_NB_soup.select("div.system_specs_top > a")
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        option = webdriver.ChromeOptions()
        option.add_argument("headless")
        service = Service(ChromeDriverManager().install())
        Lenovo_NB_data_deta = webdriver.Chrome(service=service,options=option)
        Lenovo_NB_data_deta.get("https://www.lenovo.com" + NB_deta_url[0]['href']+"#features")
        Lenovo_NB_data_deta.execute_script("document.body.style.zoom='50%'")
        
        L_NB_deta_soup = BeautifulSoup(Lenovo_NB_data_deta.page_source,'html.parser')
        NB_deta_n = L_NB_deta_soup.select("tr.item")
        j = 0
        H=[]
        W=[]
        for j in range(len(NB_deta_n)):
            NB_deta_Name = NB_deta_n[j].select("th")
            NB_deta_d = NB_deta_n[j].select("p")
            k = 0
            D = []
            for k in range(len(NB_deta_d)):
                D.append(NB_deta_d[k].text)
            D = "\n".join(D)
            if "Dimensions" in NB_deta_Name[0].text:
                Dim = D.split("/")
                D_cut = 0
                for D_cut in range(len(Dim)):
                    if len(Dim[D_cut].split("mm")) > 2:
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
                        if "W x D x H" in NB_deta_Name[0].text:
                            df_2['Height'][re_load] = De
                            df_2['Width'][re_load] = H
                            df_2['Depth'][re_load] = W
                        else:
                            df_2['Height'][re_load] = H
                            df_2['Width'][re_load] = W
                            df_2['Depth'][re_load] = De
                    elif "mm" in Dim[D_cut] and "inches" in Dim[D_cut]:
                        Dim = D.split("(")
                        hwd = 0
                        for hwd in range(len(Dim)):
                            if "mm" in Dim[hwd]:
                                H = Dim[hwd].split(":")[-1].split("x")[0].strip()
                                W = Dim[hwd].split(":")[-1].split("x")[1].strip()
                                De = Dim[hwd].split(":")[-1].split("x")[2].split("as")[-1].strip()
                            if "W x D x H" in NB_deta_Name[0].text:
                                df_2['Height'][re_load] = De
                                df_2['Width'][re_load] = H
                                df_2['Depth'][re_load] = W
                            else:
                                df_2['Height'][re_load] = H
                                df_2['Width'][re_load] = W
                                df_2['Depth'][re_load] = De
                    elif "mm" in Dim[D_cut] and (len(Dim[D_cut].split("mm")[0]) > len(Dim[D_cut].split("mm")[1])):
                        H = Dim[D_cut].split("(mm")[0].split("x")[0].strip()
                        W = Dim[D_cut].split("(mm")[0].split("x")[1].strip()
                        De = Dim[D_cut].split("(mm")[0].split("x")[2].strip()
                        if "W x D x H" in NB_deta_Name[0].text:
                            df_2['Height'][re_load] = De
                            df_2['Width'][re_load] = H
                            df_2['Depth'][re_load] = W
                        else:
                            df_2['Height'][re_load] = H
                            df_2['Width'][re_load] = W
                            df_2['Depth'][re_load] = De  
            elif "Weight" in NB_deta_Name[0].text:
                W_cut = D.split("at")[-1].split("from")[-1].split("than")[-1].split("Starting")[-1].split("g")
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
                df_2['Weight'][re_load] = Wei
                
#檢查HP_NB 的缺值/補充
re_load = 0
df_1= df_1.reset_index(drop=True)
for re_load in range(len(df_1)):
    if str(df_1['Official Price'][re_load]) == "Nan":
        print(re_load)
        Price = "Nan"
        data_url = df_1['Web Link'][re_load]
        option = webdriver.ChromeOptions()
        HP_DT_data = webdriver.Chrome(options=option)
        HP_DT_data.get(data_url)
        sleep(2)
        L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
        L_DT_data_n = L_DT_soup.select("div.Typography-module_root__eQwd4.Typography-module_boldL__LZR-5.PriceBlock-module_hasActiveDeal__W4zIr.Typography-module_responsive__iddT7")
        if len(L_DT_data_n) <1:
            L_DT_data_n = L_DT_soup.select("div.Typography-module_root__eQwd4.Typography-module_boldL__LZR-5.Typography-module_responsive__iddT7")
        HP_DT_data.quit()
        if len(L_DT_data_n) >0:           
            Price = L_DT_data_n[0].text
            Price = Price.replace("$","")
        df_1['Official Price'][re_load] = Price
        
df_1["FPR_model"] = "No"
df_2["FPR_model"] = "No"
df_1 = df_1.sort_values(["Model Name"],ascending=True)
df_2 = df_2.sort_values(["Model Name"],ascending=True)
df_1 = df_1.T   
df_2 = df_2.T
df = df.merge(df_1,how = "outer", left_index=True, right_index=True)
df = df.merge(df_2,how = "outer", left_index=True, right_index=True)

df = df.T
#Type分類
Type_NB = []
re_type_NB = 0
for re_type_NB in range(len(df["Model Name"])):
    #Dell分類
    if df.iloc[re_type_NB]["Brand"] == "Dell":
        if "chrome" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Chrome")
            df.iloc[re_type_NB]["Operating System"] = "Chrome OS"
        elif "alienware" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Gaming")
        elif "g series" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Consumer/Gaming")
        elif "inspiron" in df.iloc[re_type_NB]["Model Name"].lower() or "xps" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Consumer")
        elif "latitude" in df.iloc[re_type_NB]["Model Name"].lower() or "vostro" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Commercial")       
        elif "precision" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Workstation")
        elif "gaming"in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Consumer/Gaming")
        else:
            Type_NB.append("None Type_NB")
    #Lenovo分類        
    elif df.iloc[re_type_NB]["Brand"] == "Lenovo":
        if "ideaPad" in df.iloc[re_type_NB]["Model Name"].lower() or "slim" in df.iloc[re_type_NB]["Model Name"].lower() or "yoga" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Consumer")
        elif "legion" in df.iloc[re_type_NB]["Model Name"].lower() or "loq" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Gaming")
        elif "thinkbook" in df.iloc[re_type_NB]["Model Name"].lower() or "thinkpad" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Commercial")         
        else:
            Type_NB.append("None Type_NB")
    #Hp分類        
    elif df.iloc[re_type_NB]["Brand"] == "Hp":
        if "dragonFly pro" in df.iloc[re_type_NB]["Model Name"].lower() or "envy" in df.iloc[re_type_NB]["Model Name"].lower() or "pavilion" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Consumer")
        elif "dragonFly" in df.iloc[re_type_NB]["Model Name"].lower() or "pro" in df.iloc[re_type_NB]["Model Name"].lower() or "elite" in df.iloc[re_type_NB]["Model Name"].lower() or "fortis" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Commercial")
        elif "omen" in df.iloc[re_type_NB]["Model Name"].lower() or "victus" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Gaming")
        elif "zbook" in df.iloc[re_type_NB]["Model Name"].lower():
            Type_NB.append("Workstation")           
        else:
            Type_NB.append("None Type_NB")
df = df.T

#desktop資料重整
df1 = pd.read_excel("DELL_DT.xlsx",index_col=0)
df1 = df1.T
df1.reset_index(drop = True, inplace = True)
df1.rename(columns={'Power': 'Power Supply'}, inplace=True)
df1.rename(columns={'Storage': 'Hard Drive'}, inplace=True)

#將商品名稱與長寬深/重量抓出建立字典<方便後續補齊 減少需抓捕資料>
Dell_DT_data = {}
DNB = 0
for DNB in range(len(df1["Brand"])):
    if df1["Model Name"][DNB] not in Dell_DT_data and len(str(df1["Dimensions and Weight"][DNB])) > 25 :
        Dell_DT_data[df1["Model Name"][DNB]] = str(df1["Dimensions and Weight"][DNB])
        
Height = []
Width = []
Depth = []
Weight = []

#檢查DELL_DT缺值
#尚不需

#針對DELL重量體積資料處理 & 單位換算 & 資料切割
S_W = df1["Dimensions and Weight"]
for i in range(len(S_W)):
    if str(Dell_DT_data.get(df1["Model Name"][i])) != "None":
        A00 = Dell_DT_data[df1["Model Name"][i]].lower()
        A0 = A00.split("weight")
        no_weight = 0
        for no_weight in range(len(A0)):
            height_data = " "
            width_data = " "
            depth_data = " "
            Weight_data = " "
            if "weight" not in A0[no_weight].lower() and "kg" not in A0[no_weight].lower() and ("mm" in A0[no_weight].lower() or "cm" in A0[no_weight].lower()):
                A1_text = A0[no_weight].split("width")
                A1_besa = A0[no_weight]                   
            if len(A1_text) <= 2:
                if "mm" in A1_besa:
                    A1 = A1_besa.split("mm")
                    mm_data = 0
                    for mm_data in range(len(A1)):
                        if "height" in A1[mm_data]:
                            height_data = A1[mm_data].split("height")[-1].split("3. ")[0].split("2. ")[0].split("(")[-1].split("-")[-1].split("–")[-1].split(":")[-1].strip()
                        elif "width" in A1[mm_data]:
                            width_data = A1[mm_data].split("width")[-1].split("3. ")[0].split("2. ")[-1].split("(")[-1].split(":")[-1].strip()
                        elif "depth" in A1[mm_data] or "length" in A1[mm_data]:
                            depth_data = A1[mm_data].split("depth")[-1].split("length")[-1].split("3. ")[-1].split("(")[-1].split(":")[-1].strip()
                elif "cm" in A1_besa:
                    A1 = A1_besa.split("cm")
                    cm_data = 0
                    for cm_data in range(len(A1)):
                        if "height" in A1[cm_data]:
                            height_data = float(A1[cm_data].split("height")[-1].split("3. ")[0].split("2. ")[0].split("(")[-1].split(")")[0].split("-")[-1].split(":")[-1].strip())*10
                        elif "width" in A1[cm_data]:
                            width_data = float(A1[cm_data].split("width")[-1].split("3. ")[0].split("2. ")[-1].split("(")[-1].split(":")[-1].strip())*10
                        elif "depth" in A1[cm_data] or "length" in A1[cm_data]:
                            depth_data = float(A1[cm_data].split("depth")[-1].split("length")[-1].split("3. ")[-1].split("(")[-1].split(":")[-1].strip())*10
            if len(A0) >1:               
                if len(df1["Model Name"][i].split("Special")) > 1:
                    if "kg" in A0[-2]:
                        A2 = A0[-2].split("kg")
                        Weight_data = A2[0].split("(")[-1].split(":")[-1].strip()
                    elif "g" in A0[-2]:
                        A2 = A0[-2].split("g")
                        Weight_data = float(A2[0].split("(")[-1].split(":")[-1].strip())/1000
                else:
                    data_real_set = 0
                    for data_real_set in range(len(A0)):
                        if "g)" in A0[data_real_set] or "kg)" in A0[data_real_set]:
                            A_date = A0[data_real_set]
                            if "kg" in A_date:
                                A2 = A_date.split("kg")
                                Weight_data = A2[0].split("(")[-1].split(":")[-1].strip()
                            elif "g" in A_date:
                                A2 = A_date.split("g")
                                Weight_data = float(A2[0].split("(")[-1].split(":")[-1].strip())/1000
                            
        else:
            pass
        Weight.append(Weight_data)
        Height.append(height_data)
        Width.append(width_data)
        Depth.append(depth_data)
    else:
        Height.append("")        
        Width.append("")
        Depth.append("")
        Weight.append("")
      
#重新填充資料欄位與去除欄位
df1["Height"] = Height
df1["Width"] = Width
df1["Depth"] = Depth
df1["Weight"] = Weight
df1 = df1.drop(columns='Dimensions and Weight')
df1 = df1.sort_values(["Model Name"],ascending=True)
df1 = df1.T

#載入其他公司資料進行合併
df1_1 = pd.read_excel("HP_DT.xlsx",index_col="Unnamed: 0")    
df1_2 = pd.read_excel("Lenovo_DT.xlsx",index_col="Unnamed: 0")
df1_1 = df1_1.T
df1_2 = df1_2.T
df1_1= df1_1.reset_index(drop=True)
df1_2= df1_2.reset_index(drop=True)
HP_DT_data_Display = {}
HP_DT_data_HD = {}

#HP_DT 現有資料剪貼
DNB = 0
for DNB in range(len(df1_1["Brand"])):
    if len(str(df1_1["Display"][DNB])) < 10 and df1_1["Model Name"][DNB] in HP_DT_data_Display:
        df1_1["Display"][DNB] = HP_DT_data_Display.get(df1_1["Model Name"][DNB])
    if len(str(df1_1["Hard Drive"][DNB])) < 10 and df1_1["Model Name"][DNB] in HP_DT_data_HD:
        df1_1["Hard Drive"][DNB] = HP_DT_data_HD.get(df1_1["Model Name"][DNB])             

DNB = 0
for DNB in range(len(df1_1["Brand"])):
    if len(str(df1_1["Display"][DNB])) < 10 and df1_1["Model Name"][DNB] in HP_DT_data_Display:
        df1_1["Display"][DNB] = HP_DT_data_Display.get(df1_1["Model Name"][DNB])
    if len(str(df1_1["Hard Drive"][DNB])) < 10 and df1_1["Model Name"][DNB] in HP_DT_data_HD:
        df1_1["Hard Drive"][DNB] = HP_DT_data_HD.get(df1_1["Model Name"][DNB])
        
Lenovo_DT_data_PS = {}
Lenovo_DT_data_Display = {}

#Lenovo_DT 現有資料剪貼
DNB = 0
for DNB in range(len(df1_2["Brand"])):
    if len(str(df1_2["Ports & Slots"][DNB])) < 10 and df1_2["Model Name"][DNB] in Lenovo_DT_data_PS:
        df1_2["Ports & Slots"][DNB] = Lenovo_DT_data_PS.get(df1_2["Model Name"][DNB])
    if len(str(df1_2["Display"][DNB])) < 10 and df1_2["Model Name"][DNB] in Lenovo_DT_data_Display:
        df1_2["Display"][DNB] = Lenovo_DT_data_Display.get(df1_2["Model Name"][DNB])               

DNB = 0
for DNB in range(len(df1_2["Brand"])):
    if len(str(df1_2["Ports & Slots"][DNB])) < 10 and df1_2["Model Name"][DNB] in Lenovo_DT_data_PS:
        df1_2["Ports & Slots"][DNB] = Lenovo_DT_data_PS.get(df1_2["Model Name"][DNB])
    if len(str(df1_2["Display"][DNB])) < 10 and df1_2["Model Name"][DNB] in Lenovo_DT_data_Display:
        df1_2["Display"][DNB] = Lenovo_DT_data_Display.get(df1_2["Model Name"][DNB])

# 檢查Lenovo_DT的缺值(HWD/Weight)並補齊
re_load = 0

for re_load in range(len(df1_2)):
    if str(df1_2['Depth'][re_load]) == "nan" :
        print(re_load)
        delay = random.uniform(1.0, 5.0)
        sleep(delay)
        data_url = df1_2['Web Link'][re_load] 
        Lenovo_NB_data = requests.get(data_url + "#features", headers=my_header)
        if Lenovo_NB_data.status_code==200:
            L_NB_soup = BeautifulSoup(Lenovo_NB_data.text,'html.parser')
            #商品詳細網址            
            NB_deta_url = L_NB_soup.select("div.system_specs_top > a")
        delay = random.uniform(0.5, 5.0)
        sleep(delay)
        option = webdriver.ChromeOptions()
        option.add_argument("headless")
        service = Service(ChromeDriverManager().install())
        Lenovo_NB_data_deta = webdriver.Chrome(service=service,options=option)
        Lenovo_NB_data_deta.get("https://www.lenovo.com" + NB_deta_url[0]['href']+"#features")
        Lenovo_NB_data_deta.execute_script("document.body.style.zoom='50%'")
        
        L_NB_deta_soup = BeautifulSoup(Lenovo_NB_data_deta.page_source,'html.parser')
        NB_deta_n = L_NB_deta_soup.select("tr.item")
        j = 0
        H=[]
        W=[]
        for j in range(len(NB_deta_n)):
            NB_deta_Name = NB_deta_n[j].select("th")
            NB_deta_d = NB_deta_n[j].select("p")
            k = 0
            D = []
            for k in range(len(NB_deta_d)):
                D.append(NB_deta_d[k].text)
            D = "\n".join(D)
            if "Dimensions" in NB_deta_Name[0].text:
                Dim = D.split("/")
                D_cut = 0
                for D_cut in range(len(Dim)):
                    if len(Dim[D_cut].split("mm")) > 2:
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
                        if "W x D x H" in NB_deta_Name[0].text:
                            df1_2['Height'][re_load] = De
                            df1_2['Width'][re_load] = H
                            df1_2['Depth'][re_load] = W
                        else:
                            df1_2['Height'][re_load] = H
                            df1_2['Width'][re_load] = W
                            df1_2['Depth'][re_load] = De
                    elif "mm" in Dim[D_cut] and "inches" in Dim[D_cut]:
                        Dim = D.split("(")
                        hwd = 0
                        for hwd in range(len(Dim)):
                            if "mm" in Dim[hwd]:
                                H = Dim[hwd].split(":")[-1].split("x")[0].strip()
                                W = Dim[hwd].split(":")[-1].split("x")[1].strip()
                                De = Dim[hwd].split(":")[-1].split("x")[2].split("as")[-1].strip()
                            if "W x D x H" in NB_deta_Name[0].text:
                                df1_2['Height'][re_load] = De
                                df1_2['Width'][re_load] = H
                                df1_2['Depth'][re_load] = W
                            else:
                                df1_2['Height'][re_load] = H
                                df1_2['Width'][re_load] = W
                                df1_2['Depth'][re_load] = De
                    elif "mm" in Dim[D_cut] and (len(Dim[D_cut].split("mm")[0]) > len(Dim[D_cut].split("mm")[1])):
                        H = Dim[D_cut].split("(mm")[0].split("x")[0].strip()
                        W = Dim[D_cut].split("(mm")[0].split("x")[1].strip()
                        De = Dim[D_cut].split("(mm")[0].split("x")[2].strip()
                        if "W x D x H" in NB_deta_Name[0].text:
                            df1_2['Height'][re_load] = De
                            df1_2['Width'][re_load] = H
                            df1_2['Depth'][re_load] = W
                        else:
                            df1_2['Height'][re_load] = H
                            df1_2['Width'][re_load] = W
                            df1_2['Depth'][re_load] = De  
            elif "Weight" in NB_deta_Name[0].text:
                W_cut = D.split("at")[-1].split("from")[-1].split("than")[-1].split("Starting")[-1].split("g")
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
                            Wei = W_cut[0].split("K")[0].split("k")[0].split("/")[-1].split("(")[-1].split("<")[-1].split("From")[-1]
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
                df1_2['Weight'][re_load] = Wei

# 檢查HP_DT
re_load = 0
df1_1= df1_1.reset_index(drop=True)
for re_load in range(len(df1_1)):
    if str(df1_1['Processor'][re_load]) == "Nan" and str(df1_1['Graphics Card'][re_load]) == "Nan" and str(df1_1['Memory'][re_load]) == "Nan":
        print(re_load)
        Processor = "Nan"
        GC = "Nan"
        Memory = "Nan"
        data_url = df1_1['Web Link'][re_load]
        option = webdriver.ChromeOptions()
        HP_DT_data = webdriver.Chrome(options=option)
        HP_DT_data.get(data_url)
        sleep(2)
        L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
        L_DT_data_n = L_DT_soup.select("li.Typography-module_root__eQwd4.Typography-module_bodyS__DBLtm.Typography-module_responsive__iddT7")
        L_DT_data_n_date = L_DT_data_n[1].text
        L_DT_data_n_date = L_DT_data_n_date.split("+")
        df1_1['Processor'][re_load] = L_DT_data_n_date[0]
        df1_1['Graphics Card'][re_load] = L_DT_data_n_date[1]
        df1_1['Memory'][re_load] = L_DT_data_n_date[2]
               
    if str(df1_1['Official Price'][re_load]) == "Nan":
        print(re_load)
        Price = "Nan"
        data_url = df1_1['Web Link'][re_load]
        option = webdriver.ChromeOptions()
        HP_DT_data = webdriver.Chrome(options=option)
        HP_DT_data.get(data_url)
        sleep(2)
        L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
        L_DT_data_n = L_DT_soup.select("div.Typography-module_root__eQwd4.Typography-module_boldL__LZR-5.PriceBlock-module_hasActiveDeal__W4zIr.Typography-module_responsive__iddT7")
        if len(L_DT_data_n) <1:
            L_DT_data_n = L_DT_soup.select("div.Typography-module_root__eQwd4.Typography-module_boldL__LZR-5.Typography-module_responsive__iddT7")
        HP_DT_data.quit()
        if len(L_DT_data_n) >0:           
            Price = L_DT_data_n[0].text
            Price = Price.replace("$","")
        df1_1['Official Price'][re_load] = Price
    if str(df1_1['Height'][re_load]) == "nan" or str(df1_1['Processor'][re_load]) == "Nan" or str(df1_1['Graphics Card'][re_load]) == "Nan":
        print(re_load)
        data_url = df1_1['Web Link'][re_load]
        option = webdriver.ChromeOptions()
        HP_DT_data = webdriver.Chrome(options=option)
        HP_DT_data.get(data_url)
        sleep(2)
        L_DT_soup = BeautifulSoup(HP_DT_data.page_source,'html.parser')
        L_DT_data_n = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
        L_DT_data_d_ack = L_DT_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE")
        L_DT_specs = L_DT_soup.select("div.Container-module_root__luUPH.Container-module_container__jSUGk > div.Footnotes-module_item__LOUR3 > div")
        HP_DT_data.quit()
        specs_data = []
        W,D,H,Weight_kg = "No_Data","No_Data","No_Data","No_Data"
        if len(L_DT_specs) > 0:
            L_DT_specs_list = L_DT_specs[-1].select("span")        
            if len(L_DT_specs_list) >0:
                for specs in L_DT_specs_list:
                    specs_data.append(specs.text)
        j=0            
        for j in range(len(L_DT_data_n)):
            L_DT_data_d = L_DT_data_d_ack[j].select("div.Spec-module_valueWrapper__DTxWC > p.Typography-module_root__eQwd4.Typography-module_bodyM__XNddq.Spec-module_value__9FkNI.Typography-module_responsive__iddT7 > span")
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
            if L_DT_data_n[j].text =="Dimensions (W X D X H)":
                Dim = (N_D).split("x")
                W = round(float(Dim[0].strip())*25.4,2)
                D = round(float(Dim[1].strip())*25.4,2)
                H = round(float(Dim[2].split("in")[0].strip())*25.4,2)
            elif L_DT_data_n[j].text =="Weight":
                Weight_kg = round(float(N_D.split("lb")[0].strip())*0.4536,2)            
        df1_1['Height'][re_load] = H
        df1_1['Depth'][re_load] = D
        df1_1['Width'][re_load] = W
        df1_1['Weight'][re_load] = Weight_kg

df1_1 = df1_1.sort_values(["Model Name"],ascending=True)
df1_2 = df1_2.sort_values(["Model Name"],ascending=True)
df1_1 = df1_1.T
df1_2 = df1_2.T
df1 = df1.merge(df1_1,how = "outer", left_index=True, right_index=True)
df1 = df1.merge(df1_2,how = "outer", left_index=True, right_index=True)    
df1 = df1.T
#Type分類
Type = []
re_type_DT = 0
AA = df1["Brand"]
for re_type_DT in range(len(df1["Model Name"])):
    #Dell分類
    if df1.iloc[re_type_DT]["Brand"] == "Dell":
        if "alienware" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Gaming")
        elif "inspiron" in df1.iloc[re_type_DT]["Model Name"].lower() or "xps" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Consumer")
        elif "optiPlex" in df1.iloc[re_type_DT]["Model Name"].lower() or "vostro" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Commercial")          
        elif "precision" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Workstation")            
        else:
            Type.append("None Type")
    #Lenovo分類        
    elif df1.iloc[re_type_DT]["Brand"] == "Lenovo":
        if "ideaCentre" in df1.iloc[re_type_DT]["Model Name"].lower() or "yoga" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Consumer")
        elif "legion" in df1.iloc[re_type_DT]["Model Name"].lower() or "loq" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Gaming")
        elif "thinkbook" in df1.iloc[re_type_DT]["Model Name"].lower() or "thinkpad" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Commercial")         
        else:
            Type.append("None Type")
    #Hp分類        
    elif df1.iloc[re_type_DT]["Brand"] == "Hp":
        if "pavilion" in df1.iloc[re_type_DT]["Model Name"].lower() or "envy" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Consumer")
        elif "thin client" in df1.iloc[re_type_DT]["Model Name"].lower() or "pro" in df1.iloc[re_type_DT]["Model Name"].lower() or "elite" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Commercial")
        elif "omen" in df1.iloc[re_type_DT]["Model Name"].lower() or "victus" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Gaming")
        elif "z" in df1.iloc[re_type_DT]["Model Name"].lower():
            Type.append("Workstation")           
        else:
            Type.append("None Type")

df1 = df1.T

#docking資料重整
df2 = pd.read_excel("DELL_Dock.xlsx",index_col="Unnamed: 0")
df2 = df2.T
df2.reset_index(drop = True, inplace = True)

Dock = df2["Weight"]

#針對DELL重量進行處理 & 單位換算 & 資料切割
j = 0
for j in range(len(Dock)):
    if str(Dock[j]) != "nan":
        if "g" in str(Dock[j]):
            D_kg = str(Dock[j])
            D_kg = D_kg.split("g")[0].split("(")[-1].strip()
            df2["Weight"][j] = round(float(D_kg)/1000,2)
        elif "oz" in str(Dock[j]):
            D_kg = str(Dock[j])
            D_kg = D_kg.split("oz")[0].split("(")[-1].strip()
            df2["Weight"][j] = round(float(D_kg)*0.0283495231,2)

#重新排序<依商品名稱>   
df2 = df2.sort_values(["Model Name"],ascending=True)
df2 = df2.T

#載入其他公司資料進行合併
df2_1 = pd.read_excel("HP_Dock.xlsx",index_col="Unnamed: 0")
df2_2 = pd.read_excel("Lenovo_docking.xlsx",index_col="Unnamed: 0")
df2_1 = df2_1.T
df2_2 = df2_2.T
df2_1 = df2_1.sort_values(["Model Name"],ascending=True)       
df2_2 = df2_2.sort_values(["Model Name"],ascending=True)
df2_1 = df2_1.T
df2_2 = df2_2.T
df2 = df2.merge(df2_1,how = "outer", left_index=True, right_index=True)
df2 = df2.merge(df2_2,how = "outer", left_index=True, right_index=True)
   
#分別對合併完的資料特徵進行重新排序
df = df.T

df["Type"] = Type_NB    
#排序
df = df[["Type","Brand","Model Name","Official Price","Ports & Slots","Camera","Display","Primary Battery","Processor","Graphics Card","Hard Drive","Memory","Operating System","Audio and Speakers","Height","Width","Depth","Weight","Wireless","NFC","FPR_model","FPR",'Power Supply',"Web Link"]]
df.rename(columns={'Wireless': 'WWAN'}, inplace=True)

df1 = df1.T    
df1["Type"] = Type    
#排序
df1 = df1[["Type","Brand","Model Name","Official Price","Ports & Slots","Display","Processor","Graphics Card","Hard Drive","Memory","Operating System","Audio and Speakers","Height","Width","Depth","Weight",'Power Supply',"Web Link"]]
df.rename(columns={'Storage': 'Hard Drive'}, inplace=True)

df2 = df2.T
df2["Type"] = ""
df2 = df2[["Type","Brand","Model Name","Official Price","Ports & Slots","Power Supply","Weight","Web Link"]]

#分別對合併完的資料特徵進行名稱更換
df = df.rename(columns={'Height': 'Height(mm)','Width': 'Width(mm)','Depth': 'Depth(mm)','Weight': 'Weight(kg)'})
df1 = df1.rename(columns={'Height': 'Height(mm)','Width': 'Width(mm)','Depth': 'Depth(mm)','Weight': 'Weight(kg)'})
df2 = df2.rename(columns={'Weight': 'Weight(kg)'})

df = df.T.fillna('NA')
df1 = df1.T.fillna('NA')
df2 = df2.T.fillna('NA')
    
# #儲存資料
writer = pd.ExcelWriter('data(new)/Data_products_total.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='laptop', index=True,header = False)
df1.to_excel(writer, sheet_name='desktop', index=True,header = False)
df2.to_excel(writer, sheet_name='docking', index=True,header = False)
writer.save()
    
# except Exception as bug:
#     # 捕获并记录错误日志
#     logging.error(f"An error occurred: {str(bug)}", exc_info=True)
