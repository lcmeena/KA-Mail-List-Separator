# -*- coding: utf-8 -*-

import numpy as np
import matplotlib.pyplot as mplot

from xlwt import Workbook


#reads file, splits by line
file=open("mail2.csv", "r")
data=file.read()
#data=data.replace(,"")
data=data.split("\n")

data1=[]

newlist=[[]]
final=[]

for i in range(0,len(data)):
    if data[i] != ' USA"' and data[i] != ' Suite 204': 
        data1.append(data[i])


for i in range(0,len(data1)):
    data1[i]=data1[i].split(',"')
    

for i in range(0,len(data1)):
    for j in range(0,len(data1[i])):
        if not not data1[i][j]:
            data1[i][j]=data1[i][j].replace('"','')
            newlist.append(data1[i][j])

for i in range(1,len(newlist)):
    if newlist[i] != "USA":
        final.append(newlist[i])


names=[]
addresses=[]
cityzip=[]

for i in range(0, len(final),3):
  names.append(final[i])
  
for i in range(1, len(final),3):
    addresses.append(final[i])
  
for i in range(2, len(final),3):
    cityzip.append(final[i])


wb= Workbook()
sheet1= wb.add_sheet("KA Mail list")

for i in range(0,len(names)):
    sheet1.write(i,0,names[i])
    sheet1.write(i,1,addresses[i])
    sheet1.write(i,2,cityzip[i])



wb.save('KA Mail list.xls')





