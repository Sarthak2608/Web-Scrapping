from selenium import webdriver
options=webdriver.ChromeOptions()
def fun(s):
    c=0;
    s1=""
    for i in s:
        if(c==1):
            s1=s1+i
        if(i=='\n'):
            c=1
    return int(s1)
driver = webdriver.Chrome('C:/Users/user/data/data/chromedriver.exe')
import xlwt 
from xlwt import Workbook 
  
# Workbook is created 
wb = Workbook() 
sheet1 = wb.add_sheet('CSE') 
sheet1.write(0, 0, 'ROLL NO') 
sheet1.write(0, 1, 'NAME') 
sheet1.write(0, 12, 'YGPA') 
sheet1.write(0, 13, 'TOTAL') 
sheet1.write(0, 14, 'PERCENTAGE') 
f3=1
import time
ct=1
driver.get('http://www.bietjhs.ac.in/result2019/GetResultodd.aspx');
for j in range(1804331001,1804331020):   
    col=0
    tot=0
    col2=2
    print(j)
    option=driver.find_element_by_id('ddlSemester')
    for i in option.find_elements_by_tag_name('option'):
        if(i.get_attribute("value")=='3'):
            i.click()
    rollno=driver.find_element_by_id('txtRollNo')
    
    rollno.send_keys(j)
    submit=driver.find_element_by_id('btnSubmit')
    submit.click()
    name=driver.find_element_by_id('lblSName')
    sheet1.write(ct, col, int(j))
    col+=1
    sheet1.write(ct, col, name.text)
    col+=1
    ygpa=driver.find_element_by_id('lbloSGPA')
    table=driver.find_element_by_id('ctl04_ctl00_ctl00_grdViewSubjectMarksheet')
    tbody=table.find_element_by_tag_name('tbody')
    
    ll=1
    for k in tbody.find_elements_by_tag_name('tr'):
        ll+=1
        if(ll==1):
            continue
        else:
            c3=0
            for l in k.find_elements_by_tag_name('td'):
                c3+=1
                if(c3==1):
                    oo=0
                    for span in l.find_elements_by_tag_name('span'):
                        oo+=1
                        if(oo==2):
                            if(f3==1):
                                sheet1.write(0,col2,span.text)
                                col2+=1
                if(c3==7):
                    x=fun(l.text)
                    tot+=x
                    sheet1.write(ct, col, x)
                    col+=1

    print(name.text)
    try: 
        sheet1.write(ct, col, float(ygpa.text))
    except:
        sheet1.write(ct, col, 0)
    col+=1
    sheet1.write(ct, col, tot)
    col+=1
    sheet1.write(ct, col, float(tot/1025)*100)
    col+=1
    driver.back()
    time.sleep(1)
    ct+=1
    f3=0

# Writing to an excel  
# sheet using Python 

  
# add_sheet is used to create sheet. 
  
wb.save('EE.xls') 
