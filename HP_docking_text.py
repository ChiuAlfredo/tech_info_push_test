from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import logging
import random
from util import web_driver

# 定義要爬取的網址
url = "https://www.hp.com/us-en/shop/sitesearch?keyword=docking"
my_header = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36',
             'Cookie': 'optimizely-user-id=6ccd-f87a-0fd0-4293; aka_client_code=TW-zh; _dy_c_exps=; _dycnst=dg; _dyid=7381400044359892971; _dyid_server=7381400044359892971; crl8.fpcuid=45227714-da1e-4f5f-8f37-6ac825348491; _pnvl=false; pushly.user_puuid=Ba5aZ4kyY1yYnFQT1WND4G936j4ZdNKD; WC_PERSISTENT=RmyQFi2lvffIxOTXNl3rUuzZ%2F0ruDvCbK06Xjb38w5Y%3D%3B2023-10-06+08%3A06%3A05.572_1696579565553-44497_10151; optimizelyEndUserId=oeu1696579566697r0.7199925813939416; BVBRANDID=8bf65928-1b59-4a56-9d61-4a05ad48aada; _dy_c_att_exps=; _dy_toffset=0; hpeuck_prefs=1111; hpeuck_answ=1; _ga_vnumid=1603226F; _scid=97dedf02-d4a0-4fe8-9469-54b34be584bb; hawk_visitor_id=8d27266f-08a9-4f04-8892-b7b34c2c24d4; _rdt_uuid=1696579584985.3681dfe7-88ff-4aae-a36e-3d8bd60f69a2; _caid=c6f8bf45-1407-40ad-b321-bc070e81bcae; _fbp=fb.1.1696579585131.596500420; _tt_enable_cookie=1; _ttp=S3yvh6F5oIQwM0uWhNxVkHkPbSI; lantern=384bd082-7c39-4352-934c-571bd5a3e547; ivid=1aad2fee1ab0ddda7200ddda4223f29b4226c67411; aam_uuid=66883613473108965262498227312444713879; s_ecid=MCMID%7C62930564465370762342967807054054055955; gis-chat-call-user-uuid=n05qwy0d-iqs9-ao3h-d306-j4bdg7fzjnwv; mdLogger=false; kampyle_userid=8e9d-52cf-999b-4c5d-b085-ddd4-b914-0757; s_vnumid=108098B9; _pnlspid=32628; _abck=71B10E8F7D73C8F083851ADE077145DE~0~YAAQrV9Ky8hWbGCLAQAAd37FbAoZ3rUGRGtPrwqcIXEsy9XmyTRgK9xxH9Z+lx+LA+EUhcOqwX+9EFpznnUT1UTxbePesoOskGQYfaCD/6VpcljInkQVB9l+rD87OOEtYrjCWIkr+A2jyrXjYkg6S/cnFJoOATwgessnoE2Dk51DaqRHW4Sz8NAckIRdi/oAig2qPsdxqlL+1b2I9zIc4lxgiVl4YkTZUgTGS1GG4RvfoADwsSDNbPHQW+LsJmJSM/pl2DGQYZjnEM/xgqS2aZQx6l+Dql7zzhaAcnC6sT7HYe4Yy6U+d2Tlc9wYepEKUCanpIc/pOpiQ/H3zCj0EAcD3bt+9NMRq32rU4P6B69a7g1ykkqI8VOuuEk1m6d55zyTzbSe9ddsErgeE9s662TfZts=~-1~-1~-1; ak_bmsc=D0F7804DB0A81E191E672EF0480DB3E6~000000000000000000000000000000~YAAQrV9Ky8lWbGCLAQAAeH7FbBXuJeQ8/YNWQMSmP4PE8A2KlccB6O9Ldw5Uu0eTTTixeyp+zE3Ttc+VYw+a+yQqFNR41LGkwxd02oSokw1rkV7UhosUSvYpwjjfS1OMb7TcMkrEJ3yTpcPzbv+M3j0LIlUeeg95Cn5Gy3hlAs6fQlIwbETYMJ4PHKDGC5erQK8W/pcm7mb3fsiT2DPTDZedLKixb2X1cSZ3O1lD9448U4CRwh2/ruIeXN/3FFO3FKbMvuT8R3htkxk84rPeIsw93SmQ+ypC5SokeTzMVs+BGUsASxaND3wCbB8ve0W0As1YUK3C0WI2qj0dnD4yykHujUeyhCdGC5aqjyTvvsyN76WEYMUEXLZUgJtA7vqrjhcLYqc=; bm_sz=E8A1E85ECBBFBF0C0BA677FA0738292A~YAAQrV9Ky8pWbGCLAQAAeH7FbBV4BZUzuNF10+5RCJbomd/VDGYh2NVBsuNKyJ6KmAC73jl6f0kq8hWDo5MCYl2Q5DonPo0rHoiat3VNn39aIA1rt8JuctBM47hoa/9dtYPIPspic8KepddsGATMCyLqIb+yMptWj+6l/6vXey58L2u/iV1z2NmTYQ8r6Q9LDIo9OnfxHrpOS5mgGpFjcFwye0j9Sls02uNeA905uHlMehdCWGQuOBTdpv7tQN+98uCXcSZGp948KwCdo0GIqflH05vjTYyrTOcUlf6wdQ==~3223865~4273218; _dy_ses_load_seq=23002%3A1698336965105; _dy_csc_ses=t; _dyjsession=5415e0b65677bc0e798ba8bb7c85cddc; dy_fs_page=www.hp.com%2Fus-en%2Fshop%2Fsitesearch%3Fkeyword%3Ddocking; _dy_lu_ses=5415e0b65677bc0e798ba8bb7c85cddc%3A1698336965617; _dycst=dk.m.c.ws.; _dy_geo=TW.AS.TW_TPE.TW_TPE_Taipei; _dy_df_geo=Taiwan..Taipei; aka_client_code=TW-zh; _dy_soct=830000.1628303.1698336967; _pnss=none; WC_SESSION_ESTABLISHED=true; WC_AUTHENTICATION_-1002=-1002%2C%2F4HDXpF3hNXyjVPxYkAks6SL7SJULRQBAfPON4QntIU%3D; WC_ACTIVEPOINTER=-1%2C10151; WC_USERACTIVITY_-1002=-1002%2C10151%2Cnull%2Cnull%2Cnull%2Cnull%2Cnull%2Cnull%2Cnull%2Cnull%2C1383256606%2Cver_null%2C7Fa%2FJ3cEZzAvUbNNatmCWp8vOjKYTc%2BS2jUy6tRPJo8iwiaKNvCZKrOT9gXVgh4bGw5LkS8G6eu20uumTPJoxD2MIXxo%2BKZuK57nR5CIv5tzkg2FSK9kInnVq%2BT8BqW7Nj%2BS%2Bzoen76IACzPLKSwSEkdXzXPaePCjpvMIFKu%2Bv4tz1IisVLBhRxSXyXVrwo5gAT2MIzhbdL37T0HT40gLv9ZG3MgHmyDw9AjlD1I7E5gfTKHaQIqa6oCM8e2lfFy; WC_GENERIC_ACTIVITYDATA=[4585982463%3Atrue%3Afalse%3A0%3A7Fw5lTLakO%2BaFiNbMfmoCiKVrQcfI%2Bw%2FBinQYcfSxsI%3D][com.ibm.commerce.context.entitlement.EntitlementContext|10003%2610003%26null%26-2000%26null%26null%26null][com.ibm.commerce.context.audit.AuditContext|1696579565553-44497][com.ibm.commerce.context.globalization.GlobalizationContext|-1%26USD%26-1%26USD][com.ibm.commerce.store.facade.server.context.StoreGeoCodeContext|null%26null%26null%26null%26null%26null][com.ibm.commerce.context.experiment.ExperimentContext|null][com.ibm.commerce.context.ExternalCartContext|null][com.ibm.commerce.context.bcsversion.BusinessContextVersionContext|null][com.ibm.commerce.context.task.TaskContext|null%26null][com.ibm.commerce.giftcenter.context.GiftCenterContext|null%26null%26null][com.ibm.commerce.catalog.businesscontext.CatalogContext|10051%267000000000000000601%26false%26false%26true][com.ibm.commerce.context.preview.PreviewContext|null%26null%26false][CTXSETNAME|Store][com.ibm.commerce.context.content.ContentContext|null][com.ibm.commerce.context.base.BaseContext|10151%26-1002%26-1002%26-1]; _gcl_au=1.1.864046028.1698336968; dcm_s=1698336967810.1277891812; AMCVS_5E34123F5245B2CD0A490D45%40AdobeOrg=1; AMCV_5E34123F5245B2CD0A490D45%40AdobeOrg=-1712354808%7CMCIDTS%7C19657%7CMCMID%7C62930564465370762342967807054054055955%7CMCAAMLH-1698941767%7C11%7CMCAAMB-1698941767%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1698344167s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C4.3.0; _uetsid=f85115b0741a11eea7a8bf80dfe2ba2a; _uetvid=3e405f40641f11ee93b421be79d0a33b; OptanonConsent=isGpcEnabled=0&datestamp=Fri+Oct+27+2023+00%3A16%3A07+GMT%2B0800+(Taipei+Standard+Time)&version=202307.1.0&browserGpcFlag=0&isIABGlobal=false&hosts=&consentId=d2c0b480-aac3-45d7-8c7d-9d903a5bd20c&interactionCount=2&landingPath=NotLandingPage&groups=1%3A1%2C2%3A1%2C3%3A1%2C4%3A1%2C8%3A1&geolocation=TW%3BTPE&AwaitingReconsent=false; OptanonAlertBoxClosed=2023-10-26T16:16:07.939Z; _scid_r=97dedf02-d4a0-4fe8-9469-54b34be584bb; hawk_visit_id=2ca07956-b440-4de9-8ac9-62861f6ff2c0; kampyleUserSession=1698336968173; kampyleUserSessionsCount=2; kampyleSessionPageCounter=1; _sctr=1%7C1698336000000; CC_FINGERPRINT=1753476151; IR_gbd=hp.com; IR_5105=1698336969859%7C365159%7C1698336969859%7C%7C; JSESSIONID=0000iPFxhTOFtwvXK_fziD1iIRt:f5e95999; IR_PI=f9aadcf4-741a-11ee-b206-0bc928612d19%7C1698423369859; s_prevPage=us:en:store:public:search:shared:shared:search; s_fid=F42025EA5B3B9E02-87358B4B2D1838D0; ddj=-; s_vnum=2; s_invisit=1; _ga4_vnum=1; _ga4_vnumid=29F392D4; _ga4_invisit=1; _ga_vnum=2; s_cc=true; _cavisit=18b6cc5a40a|; _gid=GA1.2.1153211855.1698336974; _ga_KBN04WJ49D=GS1.1.1698336973.3.0.1698336973.0.0.0; _ga_invisit=6; _gat_UA-73104404-1=1; s_invisitc=2; s_engc=T; hawk_query_id=d7259855-d9b9-4cec-8d43-c116e2236c75; _ga=GA1.2.1745102095.1696579585; _ga_invisitc=2; _ga_GHQEVRZ8SH=GS1.2.1698336973.2.0.1698336974.59.0.0; _ga4_invisitc=5; RT="z=1&dm=www.hp.com&si=bb0692c2-1e72-4b04-9f61-a8ac8bd837ce&ss=lo7dz82z&sl=2&tt=-3lk&bcn=%2F%2F684d0d43.akstat.io%2F&ld=8b3"; akavpau_us_vp=1698337282~id=8e459dd6f54a4db1f88e99282e41aa84; bm_sv=CFE4E67AC47DFB75B9134438070C0631~YAAQrV9KyyRcbGCLAQAArMTFbBVlnkHCEO1zKGjpoSdaxKpsWh/i4U0ZMCpdJG5uKZwqshp9iK/3OT/ydQeJElVgCwgGNiuYOOgUrkG7ZuN4QmUdh8kHm9yxmgb1leSYeVD7haKf7T1VkruKYTtyjrlKPRf5cv4x1qU0PrCt3k1pzIOWWQrkAMbu3QrAIU5R0XsAvQxHp3AffA14brFo5KP0X6oLYAXAmEIv7hxx5GV5sCriEUdA2zqQjc8=~1'}

HP_Dock = web_driver()
HP_Dock.get(url)
HP_Dock.execute_script("document.body.style.zoom='50%'")
sleep(2)

# #頁面向下滾動 & 資料載入

# while True:
#     try:
#         #向下滾動
#         HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         sleep(2)
#         HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         #按按鈕
#         button = HP_Dock.find_elements(By.CSS_SELECTOR,'button.hawksearch-load-more')
#         HP_Dock.execute_script("arguments[0].click();", button[0])
#         sleep(2)
#         HP_Dock.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         sleep(2)
#     except:
#         break

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
for i in range(len(href_line)):
    print("HP_Dock {}".format(i))
    delay = random.uniform(2.0,6.0)
    sleep(delay)
    data_url = href_line[i]
    HP_Dock_data = web_driver()
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