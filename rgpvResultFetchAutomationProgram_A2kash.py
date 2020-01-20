from openpyxl import Workbook
from selenium import webdriver
import openpyxl
wk: Workbook = openpyxl.Workbook()
wr = openpyxl.load_workbook('insertRollNumber_A2softech.xlsx')
sh = wk['Sheet']
sh['a1'].value= "Name"
sh['b1'].value= "Roll No."
sh['c1'].value= "Course"
sh['d1'].value= "Branch"
sh['e1'].value= "Semester"
sh['f1'].value= "Result Des."
sh['r1'].value= "SGPA"
sh['s1'].value= "CGPA"
sh['g1'].value='SUB - 01'
sh['h1'].value='SUB - 02'
sh['i1'].value='SUB - 03'
sh['j1'].value='SUB - 04'
sh['k1'].value='SUB - 05'
sh['l1'].value='SUB - 06'
sh['m1'].value='SUB - 07'
sh['n1'].value='SUB - 08'
sh['o1'].value='SUB - 09'
sh['p1'].value='SUB - 10'
sh['q1'].value='SUB - 11'



driver = webdriver.Chrome(executable_path='chromedriver.exe')
import time
driver.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')                                # All Program Courses are here
driver.find_element_by_id('radlstProgram_0').click()                                            # B.E Result
#driver.find_element_by_id('radlstProgram_1').click()                                            # B.Tech Result
#driver.find_element_by_id('radlstProgram_8').click()                                            # M.Tech Result
shr = wr['Sheet1']                                                                              # CHANGE SHEET NUMBER TO FETCH ROLL NUMBER
z=1                                                                                             # INSERT DATA ON WHICH ROW : ENTER ROW NUMBER
for i in shr["a1":"a2"]:                                                                       # CHANGE CELL TO FETCH ROLL NUMBER
    for t in i:
        # time.sleep(2)
        driver.find_element_by_name('ctl00$ContentPlaceHolder1$txtrollno').send_keys(t.value)
        driver.find_element_by_name('ctl00$ContentPlaceHolder1$drpSemester').send_keys('7')    # CHANGE SEMESTER NUMBER
        time.sleep(10)
        driver.find_element_by_name('ctl00$ContentPlaceHolder1$btnviewresult').click()
        # time.sleep(1)
        z=int(z)
        z=z+1
        try:
            try:
                # name
                data = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]')
                data2 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblNameGrading')
                print(data.text, ":", data2.text)
                y='a'+str(z)
                sh[y]=data2.text
                sh['a1']=data.text

                # roll no
                data3 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]')
                data4 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblRollNoGrading')
                print(data3.text, ":", data4.text)
                y='b'+str(z)
                sh[y]=data4.text
                sh['b1'] = data3.text

                # course
                data5 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[1]')
                data6 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblProgramGrading')
                print(data5.text, ":", data6.text)
                y='c'+str(z)
                sh[y]=data6.text
                sh['c1'] = data5.text
                # branch
                data7 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[3]')
                data8 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblBranchGrading')
                print(data7.text, ":", data8.text)
                y='d'+str(z)
                sh[y]=data8.text
                sh['d1'] = data7.text
                # semester
                data9 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[2]/td/table/tbody/tr[4]/td[1]')
                data10 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblSemesterGrading')
                print(data9.text, ":", data10.text)
                y='e'+str(z)
                sh[y]=data10.text
                sh['e1'] = data9.text
                # result
                data11 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[4]/td/table/tbody/tr[1]/th[1]')
                data12 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblResultNewGrading')
                print(data11.text, ":", data12.text)
                y='f'+str(z)
                sh[y]=data12.text
                sh['f1'] = data11.text
                # SGPA
                data13 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[4]/td/table/tbody/tr[1]/th[2]')
                data14 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblSGPA')
                print(data13.text, ":", data14.text)
                y='r'+str(z)
                sh[y]=data14.text
                sh['r1'] = data13.text
                # CGPA
                data15 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[4]/td/table/tbody/tr[1]/th[3]')
                data16 = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lblcgpa')
                print(data15.text, ":", data16.text)
                y='s'+str(z)
                sh[y]=data16.text
                sh['s1'] = data15.text
            except:
                print('Information Error')
            try:
                #subject 01

                data17 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[2]/tbody/tr/td[1]')
                data18 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[2]/tbody/tr/td[4]')
                print(data17.text, ":", data18.text)
                y='g'+str(z)
                sh[y]=data18.text
                sh['g1'] = data17.text


                #subject 02
                data19 = driver.find_element_by_xpath(
                    '//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[3]/tbody/tr/td[1]')
                data20 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[3]/tbody/tr/td[4]')
                print(data19.text, ":", data20.text)
                y='h'+str(z)
                sh[y]=data20.text
                sh['h1'] = data19.text


                #subject 03
                data21 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[4]/tbody/tr/td[1]')
                data22 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[4]/tbody/tr/td[4]')
                print(data21.text, ":", data22.text)
                y='i'+str(z)
                sh[y]=data22.text
                sh['i1'] = data21.text



                #subject 04
                data23 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[5]/tbody/tr/td[1]')
                data24 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[5]/tbody/tr/td[4]')
                print(data23.text, ":", data24.text)
                y='j'+str(z)
                sh[y]=data24.text
                sh['j1'] = data23.text


                #subject 05
                data25 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[6]/tbody/tr/td[1]')
                data26 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[6]/tbody/tr/td[4]')
                print(data25.text, ":", data26.text)
                y='k'+str(z)
                sh[y]=data26.text
                sh['k1'] = data25.text


                #subject 06
                data27 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[7]/tbody/tr/td[1]')
                data28 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[7]/tbody/tr/td[4]')
                print(data27.text, ":", data28.text)
                y='l'+str(z)
                sh[y]=data28.text
                sh['l1'] = data27.text


                #subject 07
                data29 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[8]/tbody/tr/td[1]')
                data30 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[8]/tbody/tr/td[4]')
                print(data29.text, ":", data30.text)
                y='m'+str(z)
                sh[y]=data30.text
                sh['m1'] = data29.text


                #subject 08
                data31 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[9]/tbody/tr/td[1]')
                data32 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[9]/tbody/tr/td[4]')
                print(data31.text, ":", data32.text)
                y='n'+str(z)
                sh[y]=data32.text
                sh['n1'] = data31.text


                #subject 09
                data33 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[10]/tbody/tr/td[1]')
                data34 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[10]/tbody/tr/td[4]')
                print(data33.text, ":", data34.text)
                y='o'+str(z)
                sh[y]=data34.text
                sh['o1'] = data33.text


                #subject 10
                data35 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[11]/tbody/tr/td[1]')
                data36 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[11]/tbody/tr/td[4]')
                print(data35.text, ":", data36.text)
                y='p'+str(z)
                sh[y]=data36.text
                sh['p1'] = data35.text

                # subject 11

                data37 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[12]/tbody/tr/td[1]')
                data38 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlGrading"]/table/tbody/tr[3]/td/table[12]/tbody/tr/td[4]')
                print(data37.text, ":", data38.text)
                y = 'q' + str(z)
                sh[y] = data38.text
                sh['q1'] = data37.text
                # except:
                #     data37 = "N.A"
                #     data38 = "N.A"
                #     print(data37.text, ":", data38.text)
                #     y = 'q' + str(z)
                #     sh[y] = data38.text
                #     sh['q1'] = data37.text

            except:
                print("One Subject Not Available, But still continue, Don't worry")

            try:
                driver.find_element_by_name('ctl00$ContentPlaceHolder1$btnReset').click()
                wk.save('dataStorage_A2softech.xlsx')
            except:
                print("Data Storage Error : Please, Close the excel Data.xlsx File")
        except:

            element = driver.find_element_by_name('ctl00$ContentPlaceHolder1$btnReset').click()
            print('full error hai')
            continue

print("\nPlease, Open the file dataStorage_A2softech.xlsx")
print("\nCongraturation, Your Data is Created")
print("\nThank You, Visit Again")
print("\nPowered By A2softech")
driver.quit()