#! python3

# 天気ダウンロード--気象庁天気予報から秋田県の天気を取得する。

import os,requests,bs4,re,datetime,openpyxl

url='https://www.jma.go.jp/jp/yoho/309.html'




#日付を取得する
hour=str(datetime.datetime.now().day)


month=str(datetime.datetime.now().month)
day=str(datetime.datetime.now().day)


#日付が変わるまで停止する



#ページをダウンロードする
print('ページをダウンロード中{}...'.format(url))
res=requests.get(url)
res.raise_for_status()


soup=bs4.BeautifulSoup(res.text)

today_data=str(soup.select('th.weather')[0])
X=re.compile(r'title="\w+')
today_weather=X.search(today_data).group().replace('title="','')

#取得した天気予報をエクセルに書き込む
wb=openpyxl.load_workbook('C:\HP\天気自動取得\天気自動取得.xlsx')
sheet=wb.active

count_cell=1

while sheet['B'+str(count_cell)].value!=None:
    count_cell+=1

month_cell='B'+str(count_cell)
day_cell='C'+str(count_cell)
weather_cell='D'+str(count_cell)

sheet[month_cell]=month
sheet[day_cell]=day
sheet[weather_cell]=today_weather

wb.save('C:\HP\天気自動取得\天気自動取得.xlsx')






