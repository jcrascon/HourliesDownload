import sys
import requests
import datetime
import win32com.client
from win32 import win32api
from urllib.parse import urlencode
from bs4 import BeautifulSoup as bs

arguments = sys.argv
login_id = arguments[1]
password = arguments[2]
folder_path = arguments[3]

def LoginToBeta():
    login_page="https://beta.boxofficeessentials.com/login?"
    logininfo = {'hash':'','redirect_to':'','login_id':login_id,'password':password}
    with requests.Session() as r:
        r = requests.get(login_page + urlencode(logininfo))
        page_html = bs(r.text,'html.parser')
        if page_html.find("body", id="page-contents"):
                return r.cookies
        else:
                raise ValueError
try:
        cookies=LoginToBeta()
except ValueError:
        win32api.MessageBox(0, 'Password may need to be updated', 'Wrong Password', 0x00001000)
        sys.exit(1)
hourly_page = "https://beta.boxofficeessentials.com/xlsx/reports/flash/hourly_grosses_by_film"
excel_file = requests.get(hourly_page, cookies=cookies)
file_name= "Hourly_Grosses_"+datetime.datetime.now().strftime("%Y_%m_%d-%H-%M")+".xlsx"
with open(folder_path + file_name, 'wb') as output:
        output.write(excel_file.content)
        output.close

ExcelApp = win32com.client.Dispatch("Excel.Application")
wb = ExcelApp.Workbooks.Open(folder_path + file_name)
ws = wb.Sheets(1)
Rankrng = ws.Range("A6")
RankRange = ws.Range(Rankrng,Rankrng.End(4))
UniList = ['OTL','FOC','UNI','FOCW']
RankRange(len(RankRange)+1).EntireRow.Delete()
for row in range(len(RankRange),0,-1):
    distributor = ws.Range("F"+str(RankRange(row).Row)).value2
    if RankRange(row).Value>30 and distributor not in UniList:
        ws.Range("A"+str(RankRange(row).Row)).EntireRow.Delete()
wb.Close(True)
ExcelApp.Quit()