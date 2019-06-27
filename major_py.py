import serial as ss
import smtplib as sm

import xlrd as xr
import xlwt as xw
wbr=xr.open_workbook('request.xls')
shr=wbr.sheet_by_index(0)
dname=shr.col_values(0)
demail=shr.col_values(2)
dcard=shr.col_values(3)
nc=shr.ncols
r=0
def appointment():
    wbw=xw.Workbook()
    shw=wbw.add_sheet('sheet1')
    name=input('ENTER YOUR NAME:')
    dname.append(name)    
    query=input('ENTER YOUR QUERY:')
    dquery.append(query)
    email=input('ENTER YOUR EMAIL')
    demail.append(email)
    details=[dname,dquery,demail]
    for i in range(0,nc):
        for j in range(0,len(details)):
            shw.write(j,i,details[i][j])
    shw.save('requests.xls')
      
while(1):
    s=serial.Serial('/dev/tty/USB0')
    s.close()
    s.open()
    d=s.read(12)
    d=d.decode('utf-8')
    ch=-1
    if(d=='18003DB4C051'):
        ch=ch+1
        if(ch%2!=0):
            print('MAJOR IS AVAILABLE')
            sentmail(demail)
        else:
            print('MAJOR IS CURRENTLY UNAVAILABLE')
            print('TOTAL REQUESTS TII NOW- ',r)
            c=int((print('DO YOU WISH TO MAKE A REQUEST\n1.YES\n2.NO'))
            while(1):
                if(c==1):
                      appointment()
                      print('YOUR RFID IS:',dcard[r])
                      print('YOUR APPOINTMENT IS BEEN SET\nYOU WILL BE NOTIFIED AS SOON AS THE MAJOR ARRIVES')
                      r=r+1
                      break()
                else(c==0):
                      print('THANK YOU HAVE A NICE DAY')
                      break()
                elif:
                      print('WRONG CHOICE')
        
    elif(d in dcard):
        for i in range(0,len(dcard)):
            if(d==dcard[i]):
                print('WELCOME',dname[i])
    else:
        print('UNAUTHORISED ACCESS')
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
