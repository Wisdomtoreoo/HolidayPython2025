from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
import logging
import datetime
import os

# 全域變數 - add data to sharepoint
gb_site_url = "https://wisdomhk.sharepoint.com/sites/WAML"
gb_username = "Operation@wisdom-financial.com"
gb_password = "Zek65431"
# 跑原本的 HolidayPython，不能把 list name = 【HolidayPython 會錯誤】。要使用【重新命名後】的名稱 【Holiday】 ... 非常重要 
gb_list_title = "Holiday"
# gb_list_title = "Holiday_TestJudy"

# 全域變數 - share variable
gb_location = ""
gb_location_dict = {'germany': 'DE', 'hong-kong': 'HK', 'japan': 'JP', 'singapore': 'SG', 'taiwan': 'TW', 'us': 'US'}
gb_month_dict = {'JAN': '1','FEB': '2','MAR': '3','APR': '4','MAY': '5','JUN': '6','JUL': '7','AUG': '8','SEP': '9','OCT': '10','NOV': '11','DEC': '12'}
# gb_countries = ["germany", "hong-kong", "japan", "singapore", "taiwan", "us"] 
# gb_countries = ["germany", "japan", "singapore", "taiwan", "us"] 
gb_countries = ["hong-kong"] 

# logger
now = datetime.datetime.now().strftime('%Y%m%d')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

# Function to format date strings to 'm/d/yyyy' format
def format_date(date_str):
    day, month_abbr = date_str.split(' ')
    month = gb_month_dict.get(month_abbr.upper())
    if month:
        return f'{month}/{day}/{gb_year}'  # Replace 'yyyy' with the actual year
    else:
        return date_str  # Return as-is if month abbreviation is not found
    
#step 1 :
def getCountyURL(gb_year):
    logging.info(f"[getCountyURL] start")    
    for country in gb_countries:            
        url = f"https://www.timeanddate.com/holidays/{country}/{gb_year}?hol=1"    
        gb_location = gb_location_dict[country.lower()]
        #step 2 :
        df = getWebData(url,country)
        #step 3 :
        SharePointInsert(df,gb_location,gb_year)        
    logging.info(f"[getCountyURL] end")   

#step 2 :
def getWebData(url,country):
    logging.info(f"country : {country} || url : {url}")

    DaysOfWeek_list=[]; Date_list = []; Type_List = []; NameOfHoliday_List=[]

    response = requests.get(url)
    if response.status_code == 200:
        #HTML
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table', {'id': 'holidays-table'})

        #HTML Table
        table_html = str(table)
        soup = BeautifulSoup(table_html, 'html.parser')
        th_html = soup.find_all('th', class_='nw')   

        #remove tfoot element from HTML TABLE, that table clean (這裡會放置remark文字)
        tfoot_element = soup.find('tfoot')
        if tfoot_element:
            tfoot_element.extract()

        #取得-放假日期
        for YYYYMM in th_html:
            Date_list.append(YYYYMM.text.strip())

        #取得-放假名稱, 放假種類(此不用印出，用來確認是抓取national holiday)
        rows = soup.find_all('tr')
        for row in rows:
            td_elements = row.find_all('td')
            if len(td_elements) > 0:

                DaysOfWeek = BeautifulSoup(str(td_elements[0]), 'html.parser').get_text()                        
                NameOfHoliday = BeautifulSoup(str(td_elements[1]), 'html.parser').get_text()
                Type = BeautifulSoup(str(td_elements[2]), 'html.parser').get_text()

                DaysOfWeek_list.append(DaysOfWeek)
                NameOfHoliday_List.append(NameOfHoliday)            
                Type_List.append(Type)            
                
        #print(NameOfHoliday_List); print(Date_list); print(Type_List)
        #df
        data_dict = { "Date": Date_list, "DaysOfWeek_list":DaysOfWeek_list ,"NameOfHoliday": NameOfHoliday_List, "Type": Type_List }    
        df = pd.DataFrame(data_dict)

        #修改 - 日期格式
        df['Date'] = df['Date'].apply(lambda x: format_date(x))
        #留下 - Type=National holiday or TYPE=Federal Holiday or TYPE=National holiday, Christian
        df_clean_up = df[
            df['Type'].str.lower().isin(['national holiday', 'federal holiday', 'national holiday, christian'])
        ]
        #留下星期一到星期五的holiday
        df_clean_up = df_clean_up[
            df['DaysOfWeek_list'].str.lower().isin(['monday', 'tuesday', 'wednesday', 'thursday', 'friday'])
        ]

        logging.info(f"[getWebData] end")
        # print(df_clean_up)
        return df_clean_up

#step 3 :
def SharePointInsert(df,gb_location,gb_year):
    try:

        # Create an authentication context
        auth_ctx = AuthenticationContext(gb_site_url)
        auth_ctx.acquire_token_for_user(gb_username, gb_password)  # Authenticate using username and password

        # Create a SharePoint client context
        ctx = ClientContext(gb_site_url, auth_ctx)

        # Get the SharePoint list
        list_obj = ctx.web.lists.get_by_title(gb_list_title)

        df['Year'] = gb_year
        df['Location'] = gb_location

        print(df)

        # Iterate through the DataFrame and insert each row into SharePoint
        for index, row in df.iterrows():
            payload = {
                 "Title": row['NameOfHoliday']
                ,"HolidayDate": row['Date']
                ,"HolidayYear": row['Year']
                ,"Location": row['Location']
            }

            list_item = list_obj.add_item(payload)
            #Update 一筆
            #list_item.update()      
        
        #Update 全部
        list_item.update()
        ctx.execute_query()
        print(f"Data updated successfully.")
            
    except Exception as ex:
        print(f"Failed to update data. Error: {str(ex)}")        

if __name__ == "__main__":
    gb_year = "2025"
    getCountyURL(gb_year)

#---------------------------
# 核對 Log
# 1.台灣少 5/1 勞動節
# 2.香港 OK
# 3.新加坡 少 4 筆資料
#---------------------------


#用第二個 Website 去檢查輸入是否正確
#台灣
#https://publicholidays.tw/zh/2025-dates/#google_vignette
#香港
#https://www.gov.hk/tc/about/abouthk/holiday/2025.htm
#新加坡
#https://holiday.my-helper.com/zh/singapore/2025/
