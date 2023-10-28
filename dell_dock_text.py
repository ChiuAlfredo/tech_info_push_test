
from bs4 import BeautifulSoup
from time import sleep
import pandas as pd
import requests
import logging
import random

from util import web_driver
     
my_header = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36'}

#設定特徵名稱
titles = ['Type', 'Brand', 'Model Name', 'Official Price', 'Weight', 'Thunderbolt Port', 'USB-C', 'USB-A', 'Display Port', 
          'HDMI', 'LAN RJ45', 'Audio Jack', 'Power Supply', 'Web Link']
#設定網址
url = "https://www.dell.com/en-us/search/docking?p={}&c=8407%2C8408&f=true&ac=categoryfacetselect"
DELL_DOCK_data = requests.get(url, headers=my_header)
sleep(2)

i = 1
link = []
#確認連線狀況後讀取資料<連結>
if DELL_DOCK_data.status_code==200:
    L_NB_soup = BeautifulSoup(DELL_DOCK_data.text,'html.parser')
    sleep(2)
    #抓取商品數量<第一頁>
    De_dock_data = L_NB_soup.select("article")
    data = 0
    for data in De_dock_data:          
        D_deta_url = data.select("h3.ps-title > a")
        if len(D_deta_url) > 0:
            #儲存連結
            link.append("https:{}".format(D_deta_url[0]["href"]))

# #直到抓到的數量為0
# i=i+1
# while len(De_dock_data) != 0:
#     delay = random.uniform(1.0, 5.0)
#     sleep(delay)
#     url = "https://www.dell.com/en-us/search/docking?p={}&c=8407%2C8408&f=true&ac=categoryfacetselect".format(i)
#     DELL_DOCK_data = requests.get(url, headers=my_header)
#     sleep(2)
#     if DELL_DOCK_data.status_code==200:
#         L_NB_soup = BeautifulSoup(DELL_DOCK_data.text,'html.parser')
#         sleep(2)
#         #抓取特徵名稱
#         De_dock_data = L_NB_soup.select("article")
#         data = 0
#         for data in De_dock_data:          
#             D_deta_url = data.select("h3.ps-title > a")
#             if len(D_deta_url) > 0:
#                 #儲存連結
#                 link.append("https:{}".format(D_deta_url[0]["href"]))
j=0
#網頁爬取資料
for j in range(len(link)):
    delay = random.uniform(1.0, 5.0)
    sleep(delay)
    print("Dell_Dock {}".format(j))
    url_dell = link[j] + "#techspecs_section"
    Type, Brand, Model_Name, Official_Price, Weight, Thunderbolt_Port = "","Dell","","","",""
    USB_C, USB_A, Display_Port, HDMI, LAN_RJ45 = "","","","",""
    Audio_Jack, Power_Supply, Web_Link = "","",link[j]
    #開啟搜尋頁面
    
    dell_dock = web_driver()
    dell_dock.get(url_dell)
       
    dell_dock.execute_script("document.body.style.zoom='50%'")
    sleep(2)
    soup = BeautifulSoup(dell_dock.page_source,'html.parser')
    dell_dock.quit()
    
    Name = soup.select("div.pg-title > h1 > span")
    Model_Name = Name[0].text.strip()
    Money = soup.select("div.ps-dell-price.ps-simplified")
    for M in Money:
        Money_type = M.select("span")
        if len(Money_type) > 1:
            real_Money = Money_type[-1].text.split("$")[-1].strip()
            Official_Price = real_Money
        else:
            real_Money = Money[0].text.split("$")[-1].strip()
            Official_Price = real_Money
    
    dell_dock_N_D = soup.select("div.spec__child__heading")        
    dell_dock_n_D = soup.select("div.spec__child")
    k = 0       
    for dock_data in dell_dock_n_D:
        dell_dock_n = dock_data.select("div > div > div.spec__item__title")        
        dell_dock_d = dock_data.select("div.spec__item")
        l = 0
        for l in range(len(dell_dock_n)):
            if "Weight" in dell_dock_d[l].text:
                Weight = dell_dock_d[l].text.split(dell_dock_n[l].text.strip())[-1].strip()
                if "oz" in Weight:
                    oz = Weight.split("oz")[0]
                    Weight = round(float(oz)*0.0283495231,2)
                elif "kg" not in Weight and "g" in Weight:
                    g = Weight.split("g")[0]
                    Weight = round(float(g)/1000,2)
            if "Power" in dell_dock_d[l].text and "USB" not in dell_dock_d[l].text:
                Power = dell_dock_d[l].text.split(dell_dock_n[l].text.strip())[-1].strip()
                Power_Supply = Power_Supply + Power
                Power_Supply = Power_Supply.strip()
            if "Connect" in str(dell_dock_N_D[k]):
                USB_C = dell_dock_d[l].text.split(dell_dock_n[l].text.strip())[-1].strip()
        k+=1
        
    B = [Type, Brand, Model_Name, Official_Price, Weight, Thunderbolt_Port,USB_C, USB_A, Display_Port, HDMI, LAN_RJ45,Audio_Jack, Power_Supply, Web_Link]
    A = pd.Series(titles)
    if j == 0:
        C = pd.DataFrame(B,index = A)
    else:
        B = pd.DataFrame(B,index = A)
        C = C.merge(B,how = "outer",left_index=True, right_index=True)       
   
C.to_excel("DELL_Dock.xlsx")

