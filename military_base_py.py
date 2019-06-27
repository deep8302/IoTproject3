import serial 
import smtplib as sm

import xlrd as xr
import xlwt as xw
wbr=xr.open_workbook('request.xls')
shr=wbr.sheet_by_index(0)
dname=shr.col_values(0)
demail=shr.col_values(2)
dcard=shr.col_values(3)
dquery=shr.col_values(1)

nc=shr.ncols
r=0
c=0
def appointment():
    wbw=xw.Workbook()
    shw=wbw.add_sheet('sheet1')
    name=input('ENTER YOUR NAME:')    
    query=input('ENTER YOUR QUERY:')
    email=input('ENTER YOUR EMAIL')
    details=[name,query,email]
    for i in range(0,nc):
        for j in range(0,len(name)):
            shw.write(j,i,details[i][j])
    wbw.save('requests.xls')

wbr=xr.open_workbook('request.xls')
shr=wbr.sheet_by_index(0)
dname=shr.col_values(0)
demail=shr.col_values(2)
dcard=shr.col_values(3)
dquery=shr.col_values(1)

      
while(1):
    s=serial.Serial('/dev/ttyUSB0')
    s.close()
    s.open()
    d=s.read(12)
    d=d.decode('utf-8')
    ch=-1
    
    if(d=='18003DB4C051'):
        ch=ch+1
        if(ch%2!=0):
            print('MAJOR IS AVAILABLE')
            if(r!=0):
                sentmail(demail)
        else:
            print('MAJOR IS CURRENTLY UNAVAILABLE')    
    elif(d in dcard):
        for i in range(1,len(dcard)):
            if(d==dcard[i]):
                print('WELCOME')
                name=dname[i]
                print(name)
    else:
        print('UNAUTHORISED ACCESS')
        
        print('TOTAL REQUESTS TII NOW- ',r)
        c=int(input('DO YOU WISH TO MAKE A REQUEST\n1.YES\n2.NO\n'))
        while(1):
            if(c==1):
                appointment()
                print('YOUR RFID IS:',dcard[r+1])
                print('YOUR APPOINTMENT IS BEEN SET\nYOU WILL BE NOTIFIED AS SOON AS THE MAJOR ARRIVES')
                r=r+1
                break
            elif(c==2):
                print('THANK YOU HAVE A NICE DAY')
                break
            else:
                print('WRONG CHOICE')
    
s.close()


def sent_mail(ids):
    m=sm.SMTP('smtp.gmail.com',587)
    m.starttls()
    id='destrothedestroyeer@gmail.com'
    m.login(id,'destro_14')
    
    for i in range(0,2):
        message='HELLO'+str(dname[i])+'MAJOR IS AVAILABLE NOW'
        m.sendmail(id,ids[i],message)
    print('MAIL SENT')
    m.close()
