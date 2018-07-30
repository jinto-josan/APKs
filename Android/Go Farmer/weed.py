from bs4 import BeautifulSoup
import urllib.request
import time
import xlwt
from string import ascii_uppercase

generalurl="http://www.weedbusters.org.nz/weed-information/weed-list/?filterletter="
user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'
headers={'User-Agent':user_agent}
count=0
#entering the titles for worksheet
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(0, 0, "Name")
sheet1.write(0, 1, "Botanical name")
sheet1.write(0, 2, "How does it loook")
sheet1.write(0, 3, "Damage it does")
sheet1.write(0, 4, "How to get rid off")

for letter in ascii_uppercase:
    word=""

    #setting up an url
    request=urllib.request.Request(generalurl+letter,None,headers)

    html=urllib.request.urlopen(request).read()
    soup=BeautifulSoup(html,"html.parser")

    #thumb=soup.find('div',{'class':'weed-thumb'})
    #image=thumb.find('img')
    try:
        #To get the link of the weed
        web_into=soup.findAll('div',{'class':'text weed-summary'})
        for intro in web_into:
            count = count + 1
            link=intro.find('a')

        #data = soup.findAll('div', attrs = {'class ':'text weed-summary'})

        #print(image)
            name=link.string
            print (name)
            link1=link.get('href')
        #print (link.get('href'))

        #To get the page given in href of a particular weed
            request=urllib.request.Request(link1,None,headers)
            html_detail=urllib.request.urlopen(request).read()
            soup=BeautifulSoup(html_detail,"html.parser")
            data=soup.find('div',{'class':'weed-detail'})
            data1=data.findAll('p')

            sheet1.write(count, 0, link.string)
            sheet1.write(count, 1, data1[0].string)
            sheet1.write(count, 2, data1[4].string)
            sheet1.write(count, 3, data1[7].string)

            data2=data.find('ol')
            data2=data2.findAll('li')

            for i in data2:
                word=word+i.string

            sheet1.write(count, 5, i.string)
    except AttributeError as exception:
        print('error')
    except TypeError as exception:
        print('type error')
book.save("data_sheet.xls")
