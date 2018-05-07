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
for row in df.iterrows() :
    flag=False
    for r in first:
        if r.cd==row[1][1] :
            flag = True
            if row[1][7] == 1 or row[1][7] == 24 or row[1][7] == 31:
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
    if flag is False :
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
first.sort(key=lambda x: (x.laseyearrate, x.rate), reverse=True)
for r in first :
        print (str(r.cd)+'   '+r.storyname+'    '+r.area+'  '+str(r.total_budget)+'    '+str(r.performance)+'  '+str(r.laseyear_performance)+' '+r.strrate+'   '+r.strlaseyearrate)
df_answer=pd.read_excel('答案.xlsx',sheetname=0)
writer = pd.ExcelWriter('output.xlsx')
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

df1 = pd.DataFrame({'順位':emptylist,'CD': cdlist, '店名': storynamelist,'エリア名':arealist,'予算':total_budgetlist,'昨年実績':laseyear_performancelist,'実績':performancelist,'予算比':strratelist,'昨対比':strlaseyearratelist})

df1.to_excel(writer,index=False)
writer.save()
