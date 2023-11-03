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
    url = "https://www.lenovo.com/us/en/search?fq=&text=Docking&rows=60&sort=relevance&fsid=1&display_tab=Products"
    my_header = {'user-agent':UserAgent().random} 
    #開啟搜尋頁面
    Lenovo_docking = webdriver.Chrome()
    Lenovo_docking.get(url)
    Lenovo_docking.execute_script("document.body.style.zoom='50%'")
    sleep(2)
    soup_num = BeautifulSoup(Lenovo_docking.page_source,'html.parser')
    num_t = soup_num.select("p.show > span.page")
    num_n = soup_num.select("p.show > span.total")
    #頁面向下滾動 & 資料載入
    
    while num_n[0].text != num_t[0].text:
        #向下滾動
        n=0
        for n in range(3):
            Lenovo_docking.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            sleep(2)
        #按按鈕
        button = Lenovo_docking.find_elements(By.CSS_SELECTOR,'button.more')
        Lenovo_docking.execute_script("arguments[0].click();", button[0])
        sleep(2)
        soup_num = BeautifulSoup(Lenovo_docking.page_source,'html.parser')
        num_t = soup_num.select("p.show > span.page")
        Lenovo_docking.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
    
    #抓連結
    sleep(5)
    soup = BeautifulSoup(Lenovo_docking.page_source,'html.parser')
    Href = soup.select("li.product_item")
    Lenovo_docking.quit()
        
    ID_link = []
    Money_link = []
    href_line = []
    
    #將名稱與商品網址寫入list
    for link in Href:
        ID = link.select("div.product_name > a")
        Money = link.select("div.price_box > span.final-price ")
        if len(ID) > 0 and len(Money) > 0:
            ID_link.append(ID[0].text)
            new_money = Money[0].text.split("$")[-1].strip()
            Money_link.append(new_money)
            href_line.append(ID[0]['href'])
        else:
            pass
    
    #分別進入各商品頁面
    i = 0
    ip_number = 0
    new_ip = random.choice(proxy_data)[0]
    for i in range(len(href_line)):       
        delay = random.uniform(1.0, 5.0)
        sleep(delay)
        print("Lenovo_Dock {}".format(i))
        option = webdriver.ChromeOptions()
        option.add_argument("headless")
        
        #每50次換一個IP
        if ip_number == 50:
            #選擇一個代理
            new_ip = random.choice(proxy_data)[0]
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            ip_number = 0
        else:
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            
        data_url = href_line[i]      
        Lenovo_docking_data = webdriver.Chrome(options=option)
        Lenovo_docking_data.get(data_url)
        
        while "Access Denied" in Lenovo_docking_data.page_source:
            Lenovo_docking_data.quit()
            sleep(2)
            new_ip = random.choice(proxy_data)[0]
            random_proxy = new_ip
            option.add_argument("--proxy-server=http://"+random_proxy)
            ip_number = 0
            data_url = href_line[i]      
            Lenovo_docking_data = webdriver.Chrome(options=option)
            Lenovo_docking_data.get(data_url)
            
        Lenovo_docking_data.execute_script("document.body.style.zoom='50%'")
        L_docking_soup = BeautifulSoup(Lenovo_docking_data.page_source,'html.parser')
        sleep(2)
        #抓取特徵名稱 ex: CPU
        L_docking_data_n = L_docking_soup.select("tbody > tr > th")
        #抓取特徵內容
        L_docking_data_d = L_docking_soup.select("tbody > tr > td")
        #抓取特徵數量 
        L_docking_number = L_docking_soup.select("tbody > tr")
        Lenovo_docking_data.quit()
        
        ip_number = ip_number+1
        A = ["Model Name","Official Price","Web Link","Brand"]
        B = [ID_link[i],Money_link[i],data_url,"Lenovo"]
        E = ["Model Name","Official Price","Web Link","Brand"]
        Inter = []
        Thun = []
        USB = []
        Video = []
        Audio = []
        
        #處理各網頁資料
        for j in range(len(L_docking_number)):
            #篩選重複特徵
            if L_docking_data_n[j].text in E:
                pass
            else:
                #篩選需求特徵
                if L_docking_data_n[j].text == "Weight" or L_docking_data_n[j].text == "USB Ports" or L_docking_data_n[j].text == "Video Ports" or L_docking_data_n[j].text == "Thunderbolt Port" or L_docking_data_n[j].text == "Power Provided" or L_docking_data_n[j].text == "Interfaces" or L_docking_data_n[j].text == "Audio Ports": 
                    E.append(L_docking_data_n[j].text)
                    if L_docking_data_n[j].text == "Weight":
                        A.append("Weight")
                        #重量單位處理
                        if len(L_docking_data_d[j].text) > 0:
                            w = L_docking_data_d[j].text.split("g")
                            if "K" in w[0]:
                                B.append(float(w[0].split("K")[0].split(":")[-1].strip()))
                            elif "k" in w[0]:
                                B.append(float(w[0].split("k")[0].split(":")[-1].strip()))
                            elif "oz" in w[0]:
                                B.append(round(float(w[0].split("oz")[0].strip())*0.0283495231,2))
                            elif "lb" in w[0]:
                                B.append(round(float(w[0].split("lb")[0].strip())*0.4536,2)) 
                            else:
                                B.append(round(float(w[0].strip())/1000,2))
                    elif L_docking_data_n[j].text == "Power Provided":
                        A.append("Power Supply")
                        B.append(L_docking_data_d[j].text)
                    elif L_docking_data_n[j].text == "Audio Ports":
                        Audio.append(L_docking_data_d[j].text)                
                    elif L_docking_data_n[j].text == "Interfaces":
                        Inter.append(L_docking_data_d[j].text)
                    elif L_docking_data_n[j].text == "Thunderbolt Port":
                        Thun.append(L_docking_data_d[j].text)
                    elif L_docking_data_n[j].text == "USB Ports":
                        USB.append(L_docking_data_d[j].text)
                    elif L_docking_data_n[j].text == "Video Ports":
                        Video.append(L_docking_data_d[j].text)
        #port與slot合併處理                
        Ports_Slots = Inter + Thun + USB + Video + Audio
        Ports_Slots = "\n".join(Ports_Slots)                        
        A.append("Ports & Slots")
        B.append(Ports_Slots)
        #資料整合
        A = pd.Series(A)
        if i == 0:
            C = pd.DataFrame(B,index = A)
        else:
            B = pd.DataFrame(B,index = A)
            B = B[~B.index.duplicated(keep="first")]
            C = pd.merge(C,B,how = "outer",left_index=True, right_index=True)
            
    C.to_excel("Lenovo_docking.xlsx")
except Exception as bug:
    # 捕获并记录错误日志
    logging.error(f"An error occurred: {str(bug)}", exc_info=True)
