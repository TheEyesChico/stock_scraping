import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np

#pip install openpyxl

url = [
        "https://finance.yahoo.com/quote/AAPL/?guccounter=1&guce_referrer=aHR0cHM6Ly93d3cuZ29vZ2xlLmNvbS8&guce_referrer_sig=AQAAAEG1SSHaD9fUeflN7UgDsMI3YELv2yhavlacbYa3duVnUBbViHcObFPfFPve_ZnwoEZKbxHUDnOYGdfd2bXF6TwO-JRrNFDHayUMzs-4CeBvNOYXTmh1RdEIPkRlKhIR5XtncGUXjLb1AYCCtszutW6VX2XXBJr0vFF3ZqYx47ST",
        "https://finance.yahoo.com/quote/REGN?p=REGN&.tsrc=fin-srch",
        "https://finance.yahoo.com/quote/HCLTECH.NS?p=HCLTECH.NS&.tsrc=fin-srch"
        ]

def getdata(urls):
    details = {}
    for i in urls:
        r=requests.get(i)
        soup=BeautifulSoup(r.text,"html.parser")

        company=soup.find('div',{'class':'D(ib) Mt(-5px) Mend(20px) Maw(56%)--tab768 Maw(52%) Ov(h) smartphone_Maw(85%) smartphone_Mend(0px)'}).find('div',{'class':'D(ib)'}).text
        details[company]= {}

        a=soup.find('div',{'class':'D(ib) Mend(20px)'}).find_all('span')

        details[company]["Stock_Price"]=float(a[0].text)
        details[company]["Stock_Adjustment"]=a[1].text
        # b=soup.find('div',{'class':'D(ib) W(1/2) Bxz(bb) Pend(12px) Va(t) ie-7_D(i) smartphone_D(b) smartphone_W(100%) smartphone_Pend(0px) smartphone_BdY smartphone_Bdc($seperatorColor)'}).find_all('span',{'class':'Trsdu(0.3s)'})
        # print(soup.find('div',{'class':'D(ib) W(1/2) Bxz(bb) Pend(12px) Va(t) ie-7_D(i) smartphone_D(b) smartphone_W(100%) smartphone_Pend(0px) smartphone_BdY smartphone_Bdc($seperatorColor)'}).find('span',{'class':'Trsdu(0.3s)'}))
        # print(soup.find_all('div',{"class":'D(ib) Mend(20px)'})[0].find('span'))
        data_left=[]
        data_right=[]

        b=soup.find('div',{'class':"D(ib) W(1/2) Bxz(bb) Pend(12px) Va(t) ie-7_D(i) smartphone_D(b) smartphone_W(100%) smartphone_Pend(0px) smartphone_BdY smartphone_Bdc($seperatorColor)",'data-test':"left-summary-table"}).find_all('td',{'class':'Ta(end) Fw(600) Lh(14px)'})
        data_left.append(("Previous_Close", float(b[0].text)))
        data_left.append(("Open", float(b[1].text)))
        data_left.append(("Bid", b[2].text))
        data_left.append(("Ask", b[3].text))
        data_left.append(("Days_range", b[4].text))
        data_left.append(("52_week_range", b[5].text))
        data_left.append(("Volume", float(b[6].text.replace(',',""))))
        data_left.append(("Average_Volume", float(b[7].text.replace(',',""))))

        # print([print(type(i[1])) for i in data_left])
        details[company]['left_col']=data_left

        c = soup.find('div', {
            'class': "D(ib) W(1/2) Bxz(bb) Pstart(12px) Va(t) ie-7_D(i) ie-7_Pos(a) smartphone_D(b) smartphone_W(100%) smartphone_Pstart(0px) smartphone_BdB smartphone_Bdc($seperatorColor)",
            'data-test': "right-summary-table"}).findAll('td')

        data_right.append(("Market_cap", c[1].text))
        data_right.append(("Beta", float(c[3].text)))
        data_right.append(("PE_ratio", float(c[5].text.replace(',',''))))
        data_right.append(("Eps", float(c[7].text)))
        data_right.append(("Earning_date", c[9].text))
        data_right.append(("FD_yield", c[11].text))
        data_right.append(("Ex_dd", c[13].text))
        data_right.append(("1y_target", float(c[15].text.replace(',',''))))

        details[company]['right_col']=data_right
    # print(details)
    return details

details=getdata(url)

columns=["Stock_Price","Stock_Adjustment","Previous_Close","Open","Bid","Ask","Days_range","52_week_range","Volume","Average_Volume",
         "Market_cap", "Beta", "PE_ratio", "Eps", "Earning_date", "FD_yield", "Ex_dd", "1y_target"
        ]

df = pd.DataFrame(columns=columns)

for i in details:
    data=[(i[1]) for i in details[i]['left_col']] + [(i[1]) for i in details[i]['right_col']]
    basic=[details[i]['Stock_Price'], details[i]['Stock_Adjustment']]
    df.loc[i] = np.array(basic+data)

cols=["Stock_Price","Previous_Close","Open","Volume","Average_Volume",
         "Beta", "PE_ratio", "Eps", "1y_target"]

print(df.head(5))
#
# for i in cols:
#     df[i]=df[i].astype(float)

with pd.ExcelWriter('C:\\Users\\Raghu\\Desktop\\Stock_comparison.xlsx') as writer:
    df.to_excel(writer,sheet_name="hello")

# with pd.ExcelWriter('C:\\Users\\Raghu\\Desktop\\PythonExport.xlsx') as writer:
#     df.to_excel(writer,sheet_name="hello",startrow=1, header=False, index=False,startcol=1)