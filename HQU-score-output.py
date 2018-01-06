# coding=utf-8
import requests
from bs4 import BeautifulSoup
import xlwt

url = "http://10.4.12.22/server/login.aspx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36"

header = {"User-Agent": UA,
          "Referer": "http://10.4.12.22/server/"
          }

session = requests.Session()
f = session.get(url, headers=header)
soup = BeautifulSoup(f.content, "html.parser")
a = soup.find('input', {'name': '__VIEWSTATE'})['value']
b = soup.find('input', {'name': '__EVENTVALIDATION'})['value']

UserName = input("请输入用户名：")
UserPass = input("请输入密码：")

postData = {
    '__VIEWSTATE': a,
    '__EVENTVALIDATION': b,
    'UserName': UserName,
    'UserPass': UserPass,
    'ButLogin': '%E7%99%BB%E5%BD%95',
}

session.post(url, data=postData, headers=header)

f = session.get('http://10.4.12.22/server/Default.aspx', headers=header)
# print(f.content.decode())
f = session.get('http://10.4.12.22/server/Mark/StudentMark.aspx', headers=header)
# print(f.content.decode())

soup = BeautifulSoup(f.content, "html.parser")

c = soup.find('div', {'class': 'displaynone', 'id': 'Mark'})

xl = xlwt.Workbook()
sheet = xl.add_sheet('成绩单')

if c is not None:
    s = c.string
    s1 = s.split('||')
    x = 0
    y = 0
    for a in s1:
        x = 0
        s2 = a.split(':')
        for d in s2:
            sheet.write(y, x, d)
            x = x + 1
        y = y + 1
    xl.save("成绩单.xls")

else:
    print('密码输入错误')
