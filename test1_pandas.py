# coding=utf-8
import pandas as pd
class data(object) :
    yearweek=0
    cd=0
    storyname=0
    Last_year_sales_tax_withdrawal_amount=0
    Sales_tax_withdrawal_amount=0
    Execution_of_the_estimated_amount=0
    Last_year_budget_amount=0
    zone=''
    area=''
    total_budget=0
    performance=0
    laseyear_performance=0
    rate=0.0
    strrate=''
    laseyearrate=0.0
    strlaseyearrate=''
    def __init__(self):
        self.total_budget=0
df=pd.read_excel('资料.xlsx',sheetname=3)
first=[]
########核心算法开始###########
for row in df.iterrows() :#按行读取读出的数据
    flag=False #设置标志位，检测当前店铺名是否已经存在到first列表
    for r in first:#遍历first。检测当前店铺名是否已经存在到first列表
        if r.cd==row[1][1] :
            flag = True#存在置true，进行下面操作
            if row[1][7] == 1 or row[1][7] == 24 or row[1][7] == 31:#三种店铺各种值相加
                r.total_budget = r.total_budget + row[1][5]
                r.performance = r.performance + row[1][3]
                r.laseyear_performance = r.laseyear_performance + row[1][4]
                if r.total_budget != 0:
                    r.rate = r.performance / r.total_budget
                    r.strrate = '%.1f%%' % (r.rate * 100)
                else:
                    r.rate = 100
                    r.strrate = '-'
                if r.laseyear_performance != 0:
                    r.laseyearrate = r.performance / r.laseyear_performance
                    r.strlaseyearrate = '%.1f%%' % (r.laseyearrate * 100)
                else:
                    r.laseyearrate = 100
                    r.strlaseyearrate = '-'
    if flag is False :#说明first列表里没有当前店铺，新建一个data对象保存当前店铺数据
        item=data()
        item.yearweek=row[1][0]
        item.cd=row[1][1]
        item.storyname=row[1][2]
        item.area=str(row[1][10])
        item.total_budget=row[1][5]
        item.performance=row[1][3]
        item.laseyear_performance = row[1][4]
        if item.total_budget != 0 and item.area != 'nan':
            first.append(item)
first.sort(key=lambda x: (x.laseyearrate, x.rate), reverse=True)#规则排序
######核心算法结束########
for r in first :
        print (str(r.cd)+'   '+r.storyname+'    '+r.area+'  '+str(r.total_budget)+'    '+str(r.performance)+'  '+str(r.laseyear_performance)+' '+r.strrate+'   '+r.strlaseyearrate)

writer = pd.ExcelWriter('output.xlsx')
df_answer=pd.read_excel('答案.xlsx')
df1=pd.DataFrame(data={'CD':[r.cd],'店名':[r.storyname]})
df1.to_excel(writer)
cdlist=[]
storynamelist=[]
arealist=[]
total_budgetlist=[]
performancelist=[]
laseyear_performancelist=[]
strratelist=[]
strlaseyearratelist=[]
emptylist=[]
count=3
for r in first :
    cdlist.append(r.cd)
    storynamelist.append(r.storyname)
    arealist.append(r.area)
    total_budgetlist.append(float(r.total_budget)/1000)
    performancelist.append(float(r.performance)/1000)
    laseyear_performancelist.append(float(r.laseyear_performance)/1000)
    strratelist.append(r.strrate)
    strlaseyearratelist.append(r.strlaseyearrate)
    emptylist.append('')
    count=count+1

df1 = pd.DataFrame({'順位':emptylist,'CD': cdlist, '店名': storynamelist,'エリア名':arealist,'予算':total_budgetlist,'昨年実績':laseyear_performancelist,'実績':performancelist,'予算比':strratelist,'昨対比':strlaseyearratelist})
df_answer.to_excel(writer,index=False)
df1.to_excel(writer,index=False)
writer.save()
