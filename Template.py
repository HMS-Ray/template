from docxtpl import DocxTemplate
import pandas as pd
from jinja2 import Environment
jinja_env = Environment(extensions=['jinja2.ext.loopcontrols'])
import datetime
import xlrd
data = xlrd.open_workbook('信用月报.xlsm')
table = data.sheet_by_name('结果')
gm_num=int(table.cell_value(0,0))
zd_num=int(table.cell_value(0,4))
T7_num=round(table.cell_value(0,1),2)
T8_num=round(table.cell_value(0,5),2)
gm_num2=int(table.cell_value(0,8))
AA_num=int(table.cell_value(0,9))
AA_num1=int(table.cell_value(0,10))
AA_num2=gm_num2-AA_num-AA_num1
AA_num3=AA_num1+AA_num
abs_num=int(table.cell_value(0,24))
abs_num1=int(table.cell_value(0,25))
T6_num=int(table.cell_value(0,11))
gm_num4=int(table.cell_value(0,16))
sum2=round(table.cell_value(0,21),2)
gm_num3=int(table.cell_value(0,22))
zd_num2=int(table.cell_value(0,23))
gm_contents=[]
for j in range(0,gm_num):
    hh={}
    hh['name']=table.col_values(0,start_rowx=1, end_rowx=gm_num+1)[j]
    hh['cols']=table.row_values(j+1,start_colx=1,end_colx=4)
    gm_contents.append(hh)
    gm_contents[j]['cols'][0]=round(gm_contents[j]['cols'][0],2)
    gm_contents[j]['cols'][1]= "%.2f%%" % (gm_contents[j]['cols'][1] * 100)
    gm_contents[j]['cols'][2] = int(gm_contents[j]['cols'][2])
zd_contents=[]
for i in range(0,zd_num):
    hh={}
    hh['name']=table.col_values(4,start_rowx=1, end_rowx=zd_num+1)[i]
    hh['cols']=table.row_values(i+1,start_colx=5,end_colx=8)
    zd_contents.append(hh)
    zd_contents[i]['cols'][0] = round(zd_contents[i]['cols'][0], 2)
    zd_contents[i]['cols'][1] = "%.2f%%" % (zd_contents[i]['cols'][1] * 100)
    zd_contents[i]['cols'][2] = int(zd_contents[i]['cols'][2])
zb_contents=[]
for k in range(0,AA_num3):
    hh={}
    hh['name']=table.col_values(8,start_rowx=1, end_rowx=AA_num3+1)[k]
    hh['cols']=table.row_values(k+1,start_colx=9,end_colx=16)
    zb_contents.append(hh)
    zb_contents[k]['cols'][5] = round(zb_contents[k]['cols'][5], 2)
    zb_contents[k]['cols'][6] = "%.2f%%" % (zb_contents[k]['cols'][6] * 100)
absgm_contents=[]
for l in range(0,abs_num):
    hh={}
    hh['name']=table.col_values(21,start_rowx=1, end_rowx=abs_num+1)[l]
    hh['cols']=table.row_values(l+1,start_colx=22,end_colx=30)
    absgm_contents.append(hh)
    absgm_contents[l]['cols'][4] = round(absgm_contents[l]['cols'][4], 2)
    absgm_contents[l]['cols'][5] = "%.2f%%" % (absgm_contents[l]['cols'][5] * 100)
    if absgm_contents[l]['cols'][1]=='':
        absgm_contents[l]['cols'][1]='-'
abszh_contents=[]
for m in range(0,abs_num1):
    hh={}
    hh['name']=table.col_values(30,start_rowx=1, end_rowx=abs_num1+1)[m]
    hh['cols']=table.row_values(m+1,start_colx=31,end_colx=39)
    abszh_contents.append(hh)
    abszh_contents[m]['cols'][4] = round(abszh_contents[m]['cols'][4], 2)
    abszh_contents[m]['cols'][5] = "%.2f%%" % (abszh_contents[m]['cols'][5] * 100)
    if str(type(abszh_contents[m]['cols'][7]))=="<class 'float'>":
          delta=pd.Timedelta(str(int(abszh_contents[m]['cols'][7]))+'D')
          abszh_contents[m]['cols'][7]=datetime.datetime.strftime((pd.to_datetime('1899-12-30',format='%Y-%m-%d')+delta),'%Y-%m-%d')
    if abszh_contents[m]['cols'][1]=='':
        abszh_contents[m]['cols'][1]='-'
x_contents=[]
for n in range(0,gm_num4):
    hh={}
    hh['name'] = table.col_values(16, start_rowx=1, end_rowx=gm_num4 + 1)[n]
    hh['counts'] = table.row_values(n + 1, start_colx=18)[n]
    hh['bonds']=table.row_values(n + 1, start_colx=17)[n]
    hh['jz']=table.row_values(n + 1, start_colx=19)[n]
    hh['portions'] = table.row_values(n + 1, start_colx=20)[n]
    x_contents.append(hh)
    x_contents[n]['counts'] = str(int(x_contents[n]['counts']))
    x_contents[n]['jz'] = str(round(x_contents[n]['jz'], 4))
    x_contents[n]['portions'] = "%.4f%%" % (x_contents[n]['portions'] * 100)
tpl=DocxTemplate('template.docx')
context={
    'gm_contents':gm_contents,
    'zh_contents':zd_contents,
    'zb_contents':zb_contents,
    'absgm_contents':absgm_contents,
    'abszh_contents':abszh_contents,
    'gm_num':gm_num,
    'zd_num':zd_num,
    'T7_num':T7_num,
    'T8_num':T8_num,
    'gm_num2':gm_num2,
    'AA_num':AA_num,
    'AA_num1':AA_num1,
    'AA_num2':AA_num2,
    'T6_num':T6_num,
    'gm_num4':gm_num4,
    'sum2':sum2,
    'gm_num3':gm_num3,
    'zd_num2':zd_num2,
'x_contents':x_contents
}
tpl.render(context,jinja_env)
tpl.save('信用月报.docx')
