from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
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
    # 定義要爬取的網址
    url = "https://www.hp.com/us-en/shop/sitesearch?keyword=docking"
    my_header = {'user-agent':UserAgent().random}
    option = webdriver.ChromeOptions()
    
    #開啟搜尋頁面
    HP_Dock = webdriver.Chrome()
    HP_Dock.get(url)
    HP_Dock.execute_script("document.body.style.zoom='50%'")
    sleep(2)
    
    #頁面向下滾動 & 資料載入
    
    while True:
        try:
            #向下滾動
            HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            sleep(2)
            HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            #按按鈕
            button = HP_Dock.find_elements(By.CSS_SELECTOR,'button.hawksearch-load-more')
            HP_Dock.execute_script("arguments[0].click();", button[0])
            sleep(2)
            HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            sleep(2)
        except:
            break
    
    sleep(5)
    #抓連結
    soup = BeautifulSoup(HP_Dock.page_source,'html.parser')
    Href = soup.select("div.ProductTile-module_wrapper__hS8-C.productTile")
    HP_Dock.quit()
    ID_link = []
    Money_link = []
    href_line = []
    href_line_new =[]
    #將名稱與商品網址寫入list
    for link in Href:
        ID = link.select("div.ProductTile-module_details__Sgo-b > a > h3")
        Money = link.select("div.Footer-module_footer__wtYNX.ProductTile-module_footer__Gbari > div.PriceBlock-module_center__LneJE > div > div.PriceBlock-module_salePriceWrapperInline__tde7t > div > div")
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
    ip_number = 0
    new_ip = random.choice(proxy_data)[0]
    for i in range(len(href_line)):
        print("HP_Dock {}".format(i))
        delay = random.uniform(1.0, 5.0)
        sleep(delay)
        option = webdriver.ChromeOptions()
        #每50次換一個IP
        if ip_number == 50:
            new_ip = random.choice(proxy_data)[0]
            #隨機選擇一個代理
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            ip_number = 0

        data_url = href_line[i]
        HP_Dock_data = webdriver.Chrome(options=option)
        HP_Dock_data.get(data_url)
        
        while "無法連上這個網站" in HP_Dock_data.page_source:
            HP_Dock_data.quit()
            sleep(2)
            new_ip = random.choice(proxy_data)[0]
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            ip_number = 0
            data_url = href_line[i]
            HP_Dock_data = webdriver.Chrome(options=option)
            HP_Dock_data.get(data_url)
            
        HP_Dock_data.execute_script("document.body.style.zoom='50%'")
        HP_Dock_data.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        L_Dock_soup = BeautifulSoup(HP_Dock_data.page_source,'html.parser') 
        specs_data = []    
        L_D_specs = L_Dock_soup.select("div.Container-module_root__luUPH.Container-module_container__jSUGk > div.Footnotes-module_item__LOUR3 > div")
        if len(L_D_specs) > 0:
            L_D_specs_list = L_D_specs[-1].select("span.point")
            if len(L_D_specs_list) >0:
                for specs in L_D_specs_list:
                    specs_data.append(specs.text)
        #抓取特徵名稱 ex: CPU
        L_Dock_data_n = L_Dock_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerLeft__Z13zG > p")
        #抓取特徵內容
        L_Dock_data_d_ack = L_Dock_soup.select("div.Spec-module_spec__71K6S > div.Spec-module_innerRight__4TTuE > div > p.Typography-module_root__eQwd4.Typography-module_bodyM__XNddq.Spec-module_value__9FkNI.Typography-module_responsive__iddT7 > span")
        HP_Dock_data.quit()
        ip_number = ip_number + 1
        
        A = ["Model Name","Official Price","Web Link","Brand"]   
        #例外處理 搜尋頁面無顯示價格時另外加載
        if Money_link[i] =="Nan":
            new_money = L_Dock_soup.select("div.Typography-module_root__eQwd4.Typography-module_boldL__LZR-5.PriceBlock-module_hasActiveDeal__W4zIr.Typography-module_responsive__iddT7")
            nm = 0
            for nm in range(len(new_money)):
                if "$" in new_money[nm].text:
                    new_money = new_money[0].text.split("$")[-1]
                    B = [ID_link[i],new_money,data_url,"Hp"]
        else:
            B = [ID_link[i],Money_link[i],data_url,"Hp"]    
        H=[]
        W=[]
        j = 0
        port1 = []
        port2 = []
        port4 = []    
        #資料預處理
        for j in range(len(L_Dock_data_n)):
            N_D = L_Dock_data_d_ack[j].text
            if "[" in N_D and "]" in N_D:
                N_Dnum = N_D.split("[")[-1].split("]")[0]
                num = 0
                for num in range(len(N_Dnum.split(','))):
                    sp_num = 0
                    for sp_num in range(len(specs_data)):
                        if N_Dnum.split(',')[num] in specs_data[sp_num]:
                            N_D = N_D.replace(N_Dnum.split(',')[num], specs_data[sp_num].split("]")[-1])
            #篩選需求特徵
            if L_Dock_data_n[j].text == "External Ports 01":
                port1 = N_D
            elif L_Dock_data_n[j].text == "External Ports 02":
                port2 = N_D
            elif L_Dock_data_n[j].text == "External Ports 04":
                port4 = N_D
            elif L_Dock_data_n[j].text == "Power supply":
                A.append("Power Supply")
                B.append(N_D.split("[")[0])
            elif L_Dock_data_n[j].text == "Weight":
                A.append("Weight")
                Weight_kg = float(N_D.split("lb")[0].strip())*0.4536
                B.append(round(Weight_kg,2))           
        #針對port與slot複合合併處理        
        if len(port1) > 0:
            if  len(port2) > 0:
                if len(port4) > 0:
                    port = "{}\n{}\n{}".format(port1,port2,port4)
                else:
                    port = "{}\n{}".format(port1,port2)
            else:
                port = port1
            A.append("Ports & Slots")
            B.append(port)        
        #資料合併    
        A = pd.Series(A)
        if i == 0:
            C = pd.DataFrame(B,index = A)
        else:
            B = pd.DataFrame(B,index = A)
            C = C.merge(B,how = "outer",left_index=True, right_index=True)
            C.drop_duplicates(inplace=True)
    
    C.to_excel("HP_Dock.xlsx")
except Exception as bug:
    # 捕获并记录错误日志
    logging.error(f"An error occurred: {str(bug)}", exc_info=True)