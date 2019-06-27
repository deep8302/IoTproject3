import xlrd as xr
import xlwt as xw

r=0
c=0

def excel(r):
    
    wbr=xr.open_workbook('requests.xls')
    shr=wbr.sheet_by_index(0)
    dname=shr.col_values(0)
    dquery=shr.col_values(1)
    demail=shr.col_values(2)
    dcard=shr.col_values(3)
    nc=shr.ncols
    
    wbw=xw.Workbook()


    shw=wbw.add_sheet('sheet1',cell_overwrite_ok=True)
    name=input('ENTER YOUR NAME:')
    dname.append(name)
    query=input('ENTER YOUR QUERY:')
    dquery.append(query)
    email=input('ENTER YOUR EMAIL')
    demail.append(email)
    #print('SWIPE YOUR CARD')
    #card=s.read(12)
    #card=card.decode('utf-8')
    card=input('ENTER CARD NUMBER:')
    dcard.append(card)
    details=[dname,dquery,demail,dcard]
    for i in range(0,nc):
        for j in range(0,len(dname)):
            shw.write(j,i,details[i][j])
    wbw.save('requests.xls')
    return (card)
while(1):
    r=r+1
    id=excel(r)
    print('id=',id)



