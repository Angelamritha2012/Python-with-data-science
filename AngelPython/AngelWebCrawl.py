#Python Project done By ANGEL AMRITHA NIIT MADURAI BP ROAD CENTRE  
#key modules py -m pip install pyexcel-xls
#key modules py -m pip install beautifulsoup4
#key modules py -m pip install xlsxwriter
from pyexcel_xls import save_data
from pyexcel_xls import read_data
from urllib import *

from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter
import json
import socket

data=read_data("input.xls")  
ss=json.dumps(data)
object = json.loads(ss)
urlNames=dict()
wb=xlsxwriter.Workbook('output_scrap.xlsx')  
h1=wb.add_format({'bold':True,'font_color':'red'})
h2=wb.add_format({'bold':True,'font_color':'blue'})
h3=wb.add_format({'bold':True,'font_color':'green'})


for o in object:
    ws=wb.add_worksheet(str(o))
    ws.write("A1",object[o][0][0],h1)
    ws.write("A3","WORDS",h2)
    ws.write("B3","COUNT",h2)
    chart=wb.add_chart({'type':'column'})
    req=urllib.request.Request(str(object[o][0][0]),data=None,headers={'User-Agent':'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'})
    f=urllib.request.urlopen(req)
    s=f.read().decode('utf-8')
    soup=BeautifulSoup(s,"html.parser")
    socket.getaddrinfo('localhost',8080)
    for script in soup(["script","style"]):
        script.extract()
    t=soup.get_text()
    text = "".join([s for s in t.splitlines(True) if s.strip("\r\n")])
    L=text.split()
    wor=[]
    for i in range(1,len(object[o])):
        wor.append(object[o][i][0])
           
    words=sorted(set(wor))
    countD=dict()
    i=3
    for w in words:
       countD[w]=L.count(w)
       ws.write(i,0,w,h3)
       ws.write(i,1,L.count(w),h3)
       i+=1
    area="="+str(o)+"!$B$4:$B$"+str(i)     
    cate="="+str(o)+"!$A$4:$A$"+str(i)
    chart.add_series({'categories':cate ,'values': area})
    chart.set_title({'name': 'WORD COUNT'})
    chart.set_x_axis({'name': 'Words'})
    chart.set_y_axis({'name': 'Count'})
    chart.set_style(37)
    ws.insert_chart("F4",chart)
       
    print(countD)
    
    print("---------------------------------------------------------------------")
    
wb.close()






