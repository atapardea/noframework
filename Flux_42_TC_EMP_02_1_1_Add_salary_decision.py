#=======================================================================================================================
#
#       42.UC – SYS.01.Add salary decision
#
#         Steps:
#  1. User opens page
#  2. User logs in the application
#  3. User accesses the Employees record > ACTIVE Employment decisions & agreements menu
#  4. User selects Authority from the Current entity menu
#  5. User selects "Persons with incomplete or ended employment decisions/agreements" and presses Find
#  6. User selects entity from "List all Persons" and presses the "View" button
#  7. User selects entity form "List all Employments" and preses the "View" button
#  8. User scrolls to the Decisions tab on the bottom of the page and presses the "Add" button
#  9. USer selects "Change of salary" from the pop-up menu and presses Select
# 10. User adds data to the mandatory fields
# 11. User presses the "Save" button
# 12. User presses the "List" button
#
#
#   Expected result: User can view the newly created salary decision in the Decision tab
#
#   Created on: 23.03.2016
#
#=======================================================================================================================
#Imports

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import platform


from time import sleep
import sys
import traceback
import os
from selenium.webdriver.common import alert
from Utils import Utils
import time
from datetime import datetime
from selenium import webdriver



from TestPackage_NOU.Config import ConfigurationData
from TestPackage_NOU.Tools import Random


sheetCredentiale  ="Credentiale_App"





def run_Flux42_TC_EMP_02_1_1_Add_salary_decision(credentiale, fisierInput, sheetIn, fisierOutput, chromePath):

    testConfig = ConfigurationData()
    nrIteratii = testConfig.get_value_for("iterations")
    SLEEP_TIME = testConfig.get_value_for("sleep-time")
    pas = 0 ## se defineste un cursor de scriere in fisierul de rezultate pe linie, statusul fiecare actiuni

    try:
        for i in range(0,int(nrIteratii)):


            sheetOut = "Add salary decision" +"_"+ str(i+1) ## denumire Sheet cu detalii pe statusul pt fiecare actiune


        # Citeste din fisierul Config.xlsx, sheet Credentiale_App, pe coloana, valorile parametrilor
            APPLICATION_ADDRESS = Utils.read_excel(credentiale, 'Credentiale_App', 2, 1)
            USERNAME = Utils.read_excel(credentiale, 'Credentiale_App', 2, 2)
            PASSWORD = Utils.read_excel(credentiale, 'Credentiale_App', 2, 3)
            authority = Utils.read_excel(credentiale, 'Flux_22', 2, 1)


        #testconfig - PATHS - XML
            #PAS 2
            usernameLogin=testConfig.get_value_for ("id_usernameLogin")
            passwordLogin=testConfig.get_value_for("id_passwordLogin")
            loginButton=testConfig.get_value_for("id_loginButton")
            xpathValid2=testConfig.get_value_for("xpath_validatorpas2")

            #PAS 4
            xpathPinButton=testConfig.get_value_for("xpath_authority")
            idCampAuthority=testConfig.get_value_for("id_campauthority")
            xpathAuthTable=testConfig.get_value_for("xpath_authtable")
            xpathFoundAuthority=testConfig.get_value_for("xpath_foundauthority")
            xpathSaveAuthority=testConfig.get_value_for("xpath_saveauthority")
            idAuthorityField=testConfig.get_value_for("id_authorityfield")

            #PAS 5
            xpathPersonsTable=testConfig.get_value_for("xpath_personstable")
            xpathThirdRadioButton=testConfig.get_value_for("xpath_thirdradio")
            idSubmitButton=testConfig.get_value_for("id_submitbutton")

            #PAS 6
            xpathFirstPerson=testConfig.get_value_for("xpath_firstperson")
            xpathViewPerson=testConfig.get_value_for("xpath_viewperson")
            validatorPas6=testConfig.get_value_for("id_validatorPas6")

            #PAS 7
            xpathFirstEmployment=testConfig.get_value_for("xpath_firstemployment")
            xpathEmploymentView=testConfig.get_value_for("xpath_employmentview")
            validatorPas7=testConfig.get_value_for("xpath_validatorpas7")

            #PAS 8
            idBottomTable=testConfig.get_value_for("id_bottomtable")
            idAddDecision=testConfig.get_value_for("id_adddecision")
            idValidator8=testConfig.get_value_for("id_validator8")

            #PAS 9
            xpathChangeOfSalary=testConfig.get_value_for("xpath_changeofsalary")
            allSelectButtons=testConfig.get_value_for("xpath_allseelctbuttons")
            xpathValid9=testConfig.get_value_for("xpath_validpas9")

            #PAS 10
            idSalaryModificationDate=testConfig.get_value_for("id_salarymodificationdate")
            idSalaryModificationNumber=testConfig.get_value_for("id_salarymodificationnumber")
            xpathSalaryButton=testConfig.get_value_for("xpath_salaryButton")
            xpathFirstSalaryType=testConfig.get_value_for("xpath_firstsalarytype")
            xpathSelectButtons=testConfig.get_value_for("xpath_selectButtons")
            xpathSalaryValue=testConfig.get_value_for("xpath_salaryvalue")
            xpathPrecentageValue=testConfig.get_value_for("xpath_precentagevalue")
            idStartDate=testConfig.get_value_for("id_StartDate")

            #PAS 11
            xpathSaveData=testConfig.get_value_for("xpath_savedata")

            #PAS 12
            xpathListButton=testConfig.get_value_for("xpath_listbutton")
            xpathDecisionsList=testConfig.get_value_for("xpath_decisionslist")




            #=====Rezultate in fisier
            pas+=1
            Utils.write_excel_results(fisierOutput,sheetOut,pas,"Nr_TC", "Descriere", "Status", "Mentiuni")
            ## Scrie header-ul din fisierul de rezultate

#------------------------------------------------------------------
#1. User opens page
#------------------------------------------------------------------

            try:
                browser = webdriver.Chrome(chromePath)
                browser.maximize_window()
                browser.get(APPLICATION_ADDRESS)
                wait = WebDriverWait(browser, 20)

            #validate if the login page appears
                try:
                    validatorPas1=browser.find_element_by_id("loginForm")
                    if validatorPas1.is_displayed():
                        pas+=1
                        Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Deschidere pagina login", "PASS", "Pagina de login a fost accesata cu succes")
                except:
                    pas+=1
                    Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Deschidere pagina login nereusita", "FAIL","Pagina de login nu a putut fi accesata")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Deschidere browser ", "N/A","Nu a fost deschis browser-ul")
                raise Exception("Nu a fost deschis browser-ul")
#------------------------------------------------------------------
#2. User logs in the application
#------------------------------------------------------------------


        #Username
            try:
                sleep(1)
                wait.until(EC.presence_of_element_located((By.ID,usernameLogin))).send_keys(USERNAME)
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Completare camp username", "PASS", "Campul Username a fost completat cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Completare camp username", "N/A","Calea catre input-ul de utilizator nu este corecta")
                raise Exception("Adaugare username nereusita")

        #Password
            try:
                sleep(1)
                wait.until(EC.presence_of_element_located((By.ID,passwordLogin))).send_keys(PASSWORD)
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1),"Completare camp parola","PASS","Campul Parola a fost completat cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Completare camp parola", "N/A","Calea catre input-ul de parola nu este corecta")
                raise Exception("Adaugare parola nereusita")

        #Buton Login
            try:
                sleep(1)
                wait.until(EC.presence_of_element_located((By.ID,loginButton))).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1),"Apasare Login","PASS","Butonul Login a fost apasat cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Completare camp parola", "N/A","Calea catre input-ul de parola nu este corecta")
                raise Exception("Incercare de Login nereusita")


        #validate if the user was succesfully logged in

            try:
                validatorPas2=wait.until(EC.presence_of_element_located((By.XPATH,xpathValid2))).text
                if ("welcome to application!")in validatorPas2 :
                        pas+=1
                        Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare Login", "PASS", "Operatiunea de login s-a facut cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare login", "FAIL","Operatiunea de login nereusita")
                raise Exception ("Login nereusit")
#------------------------------------------------------------------
#3. User accesses the Employees record > ACTIVE Employment decisions & agreements menu
#------------------------------------------------------------------
            try:
            #access said manu
                wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,"Employees record"))).click()
                sleep(0.2)
                wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,"ACTIVE Employment decisions & agreements"))).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Accesare pagina ACTIVE Employment decisions & agreements", "PASS", "Accesare pagina ACTIVE Employment decisions & agreements reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Accesare pagina Authority ", "N/A", "Accesarea paginii de autoritati nereusita")
                raise Exception("Accesare paginii Autority nereusita")

            #Validate that the menu was accessed
            if "servicereports" in browser.current_url:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare accesare meniu ACTIVE Emp decisions", "PASS", "Validarea accesarii a meniului ACTIVE Employment decisions & agreement s-a facut cu succes")
            else :
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare accesare meniu ACTIVE Emp decisions", "FAIL","Validare accesare meniu ACTIVE Emp decisions nereusita")
                raise Exception("Validarea de acces a meniului ACTIVE Employment decisions & agreements menu nereusita")


#------------------------------------------------------------------
#4. User selects Authority from the Current entity menu
#------------------------------------------------------------------
            try:
                #save current window hande in order to come back to it later
                currentframe=browser.current_window_handle
                sleep(0.2)

                #change frame by clicking on authority menu
                wait.until(EC.presence_of_element_located((By.ID,"currentoename"))).click()
                sleep(0.2)
                browser.switch_to_frame("currentOEIFrame")
                sleep(1)
                browser.find_element_by_xpath(xpathPinButton).click()

                #press the filter button and press Search after authority
                wait.until(EC.presence_of_element_located((By.ID,idCampAuthority))).send_keys(authority)
                sleep(0.2)
                browser.find_element_by_id(idCampAuthority).send_keys(Keys.ENTER)

                #algorithm to wait untill authority is found
                tableAuthorities=browser.find_elements_by_xpath(xpathAuthTable)
                maxTableAuthorities=len(tableAuthorities)
                failsafe=0

                while len(tableAuthorities) == maxTableAuthorities:
                    sleep(1)
                    tableAuthorities=browser.find_elements_by_xpath(xpathAuthTable)
                    failsafe +=1
                    if failsafe==60:
                        raise Exception("Force exit")

                #click on the found authority
                browser.find_element_by_xpath(xpathFoundAuthority).click()

                #go back to previous frame/window and click on the Save button
                browser.switch_to_window(currentframe)
                sleep(1)
                wait.until(EC.presence_of_element_located((By.XPATH,xpathSaveAuthority))).click()

                #algorithm to wait untill the authority is selected
                searchAuthorityBox=browser.find_element_by_id(idAuthorityField)
                sleep(1)
                failsafe=0
                while str(authority) not in str(searchAuthorityBox.get_attribute("value")):
                    sleep(1)
                    searchAuthorityBox=browser.find_element_by_id(idAuthorityField)
                    failsafe+=1
                    if failsafe ==60:
                        raise Exception("Exiting infinite loop")

                #log that everything is OK
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare autoritate", "PASS", "Selectarea autoritatii s-a facut cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare autoritate", "N/A", "Selectarea autoritatii nereusita")
                raise Exception("Selectarea autoritatii nu s-a realizat")

            #VALIDATE that the right authority was selected
            if str(authority) in  str(searchAuthorityBox.get_attribute("value")):
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare autoritate", "PASS", "Validarea autoritatii s-a facut cu succes")
            else:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare autoritate", "FAIL", "Validarea autoritatii nereusita")

#------------------------------------------------------------------
#5. User selects "Persons with incomplete or ended employment decisions/agreements" and presses Find
#------------------------------------------------------------------
                #click on radio button and press Find
            try:
                wait.until(EC.presence_of_element_located((By.XPATH,xpathThirdRadioButton))).click()
                sleep(0.2)
                wait.until(EC.presence_of_element_located((By.ID,idSubmitButton))).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Cautare Persoane", "PASS", "Cautarea persoane realizata cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Cautare persoane", "N/A", "Cautarea de persoane nereusita")
                raise Exception("Cautarea de persoane nu s-a realizat cu succes")


            #VALIDATE that the search happened
            personsList=browser.find_elements_by_xpath(xpathPersonsTable)

            maxPersonsList=len(personsList)
            failsafe=0

            while maxPersonsList==len(personsList):
                sleep(1)
                personsList=browser.find_elements_by_xpath(xpathPersonsTable)
                failsafe+=1
                if failsafe==30:
                    pas+=1
                    Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare cautare in tabel persons", "FAIL", "Validarea cautarii in tabelul persons nereusita")
                    raise Exception("Exiting infinite loop - search persons table")
            pas+=1
            Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare cautare in tabel persons", "PASS", "Validarea cautarii in tabelul persons realizata cu succes")

#------------------------------------------------------------------
#  6. User selects entity from "List all Persons" and presses the "View" button
#------------------------------------------------------------------

            #click on the first element in the list and press View
            try:
                wait.until(EC.presence_of_element_located((By.XPATH,xpathFirstPerson))).click()
                sleep(0.2)
                browser.find_element_by_id(xpathViewPerson).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare persoana si apasare View", "PASS", "Selectarea persoanei si apasarea butonului View s-a realizat cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare persoana si apasare View", "N/A", "Selectarea persoanei si apasarea butonului View nu a reusit")
                raise Exception("View person action could not be done")

            #Validate that the View button was pressed
            try:
                wait.until(EC.presence_of_element_located((By.ID,validatorPas6)))
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare Selectare persoana si apasare View", "PASS", "Validare selectarii persoanei si apasarii View reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare Selectare persoana si apasare View", "FAIL", "Validare selectarii persoanei si apasarii View nereusita")
                raise Exception("Validarea apsarii View nereusita")

#------------------------------------------------------------------
#7. User selects entity from "List all Employments" and preses the "View" button
#------------------------------------------------------------------

            try:
                sleep(0.5)
                # WAIT UNTILL "LOADING" IS COMPLETED
                atp=browser.find_element_by_xpath("//div[contains(text(),'Loading...')]")
                failsafe=0

                while atp.is_displayed():
                    sleep(1)
                    failsafe+=1
                    if failsafe ==100:
                        raise Exception("Exit infinite loop")
                sleep(1)
                #select entity from the employments table
                wait.until(EC.presence_of_element_located((By.XPATH,xpathFirstEmployment))).click()
                sleep(0.5)
                #Press the View button
                wait.until(EC.presence_of_element_located((By.XPATH,xpathEmploymentView))).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare employment", "PASS", "Selectare Employment realizata cu succes")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare employment", "N/A", "Selectare Employment nerealizata")

            #Validate that the employment was selected and view button was pressed
            try:
                wait.until(EC.presence_of_element_located((By.XPATH,validatorPas7)))
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare Selectare employment si apasare View", "PASS", "Validare Selectare employment si apasare View reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare Selectare employment si apasare View", "FAIL", "Validare Selectare employment si apasare View nereusita")
                raise Exception("Validation on selecting Employment failed")
#------------------------------------------------------------------
#8. User scrolls to the Decisions tab on the bottom of the page and presses the "Add" button
#------------------------------------------------------------------
            try:
                #Go to decision table on the bottom of the page
                elem=wait.until(EC.presence_of_element_located((By.ID,idBottomTable))).location
                scriptGoBottom="window.scrollTo(0,"+str(int(elem["y"]))+");"
                browser.execute_script(scriptGoBottom)

            #Press the Add button
                browser.find_element_by_id(idAddDecision).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare Add decision", "PASS", "Apasare Add decision reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare Add decision", "N/A", "Apasare Add decision nereusita")
                raise Exception ("Add button not pressed")

            #Validate that the ADD button was pressed
            try:
                wait.until(EC.presence_of_element_located((By.ID,idValidator8)))
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare Add decision", "PASS", "Validare apasare Add decision reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare Add decision", "FAIL", "Validare apasare Add decision nereusita")
                raise Exception("Validare apasare Add button nereusita")

#------------------------------------------------------------------
#9. USer selects "Change of salary" from the pop-up menu and presses Select
#------------------------------------------------------------------

            try:
                #select "Change of salary"
                wait.until(EC.presence_of_element_located((By.XPATH,xpathChangeOfSalary))).click()
                sleep(0.5)
                #click Select (Select element code is not unique, get all elements in a list and select the one that is visible)
                selectList=browser.find_elements_by_xpath(allSelectButtons)
                for i in range (len(selectList)):
                    if selectList[i].is_displayed():
                        selectList[i].click()
                        break
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare decizie", "PASS", "Selectare decizie reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Selectare decizie", "N/A", "Selectare decizie nereusita")
                raise Exception ("Selecting decision type failed")

            #Validare selectare decizie
            validatorpas9=wait.until(EC.presence_of_element_located((By.XPATH,xpathValid9)))

            if "Change of salary" in str(validatorpas9.text):
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare selectare decizie", "PASS", "Validare selectare decizie reusita")
            else:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare selectare decizie", "FAIL", "Validare decizie nereusita")
                raise Exception("Validare selectare decizie nereusita")

#------------------------------------------------------------------
# 10. User adds data to the mandatory fields
#------------------------------------------------------------------
            try:
                #add Salary modification decision date*:
                decisiontime=time.strftime("%d-%m-%Y")
                wait.until(EC.presence_of_element_located((By.ID,idSalaryModificationDate))).send_keys(decisiontime)
                sleep(0.2)

                #add Salary modification decision number*:
                decisionnumber=Random.random_number(6)
                browser.find_element_by_id(idSalaryModificationNumber).send_keys(decisionnumber)
                sleep(0.2)

                #select type of salary
                browser.find_element_by_xpath(xpathSalaryButton).click()
                wait.until(EC.presence_of_element_located((By.XPATH,xpathFirstSalaryType))).click()
                sleep(0.2)
                selectList=browser.find_elements_by_xpath(xpathSelectButtons)
                for i in range (1,len(selectList)+1):
                    if selectList[i].is_displayed():
                        selectList[i].click()
                        break

                #select salary value
                wait.until(EC.presence_of_element_located((By.XPATH,xpathSalaryValue))).send_keys(Random.random_number(6))
                sleep(0.2)

                #add precentage value
                browser.find_element_by_xpath(xpathPrecentageValue).send_keys(Random.random_number(6))

                #add Start date of salary modification*:
                browser.find_element_by_id(idStartDate).send_keys(decisiontime)
                sleep(0.2)
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Adaugare date in campuri", "PASS", "Adaugare date in campuri reusita reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Adaugare date in campuri", "N/A", "Adaugare date in campuri nereusita")
                raise Exception("Adaugare date in campuri nereusita")



#------------------------------------------------------------------
#11. User presses the "Save" button
#------------------------------------------------------------------
            try:
                #Press the Save button
                browser.find_element_by_id(xpathSaveData).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare Save button", "PASS", "Apasare Save button reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare Save button", "PASS", "Apasare Save button nereusita")
                raise Exception("Apasarea butonului Save nereusita")

            #validate if the save button was pressed
            try:
                wait.until(EC.alert_is_present()).accept()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare Save button", "PASS", "Validare apasare Save button reusista")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare Save button", "FAIL", "Validare apasare Save button nereusista")
                raise Exception("Butonul Save nu a fost apasat")
#------------------------------------------------------------------
#12. User presses the "List" button
#------------------------------------------------------------------
            try:
                #Press the list button
                wait.until(EC.presence_of_element_located((By.XPATH,xpathListButton))).click()
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare List button", "PASS", "Apasare List button reusita")
            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Apasare List button", "N/A", "Apasare List button nereusita")
                raise Exception("Butonul List nu a fost apasat")

            # FINAL VALIDATION
            #validate that the list button was pressed AND that the element was created
            try:
                sleep(1)
                decisionsList=wait.until(EC.presence_of_all_elements_located((By.XPATH,xpathDecisionsList)))
                sleep(1)
                for i in range (2,len(decisionsList)+1):
                    xpathOfChild=str("//*[@id='grid_l_ro_teamnet_hrmisforall_domain_rmc_decision']/tbody/tr["+str(i)+"]/td[5]")
                    validatorpas12=browser.find_element_by_xpath(xpathOfChild).text
                    if str(decisionnumber)==validatorpas12:
                        pas+=1
                        Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare List button", "PASS", "Validare apasare List button reusita")
                        break

            except:
                pas+=1
                Utils.write_excel_results(fisierOutput,sheetOut,pas, "TC"+str(pas-1), "Validare apasare List button", "FAIL", "Validare apasare List button nereusita")
                raise Exception("Validare finala nereusita")
#------------------------------------------------------------------
#------------------------------------------------------------------

    except Exception as e:
        print(e)
        exc_info = sys.exc_info()
        x = traceback.print_exception(*exc_info)
        del exc_info
        pas = pas+1
        Utils.write_excel_results(fisierOutput,sheetOut, pas+1, "","Pas "+str(pas-2),"Detalii Eroare: ""N/A",str(e) + " " + str(x))



    finally:
        try:
            try:
                browser.close()
                browser.quit()
                print("Browserul a fost inchis!")
            except:

             if "Windows" in platform.platform():
                 os.system("taskkill /im chromedriver.exe /f")
                 time.sleep(1)
                 os.system("taskkill /im chrome.exe /f")
             time.sleep(4)
        except:
            print('Chrome driver este inchis')




if __name__ == '__main__':

    #=====Variabile
    credentiale= "Config.xlsx"
    fisierInput = "DateIntrare.xlsx"
    sheetIn = "Credentiale_App"
    fisierOutput = "Rezultate.xlsx"
    chromePath = 'C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe'
    timestamp = datetime.now().strftime("%d/%m/%y %H:%M:%S")

    if not os.path.exists(os.getcwd() + os.path.sep+"Logs"+os.path.sep):
        os.mkdir(os.getcwd() + os.path.sep+"Logs")

    save_folder = os.getcwd() + os.path.sep+"Logs"+os.path.sep+ str(timestamp).replace('/','.').replace(':','.')
    os.mkdir(save_folder)

    fisierOutput = save_folder+ os.path.sep+"Rezultate_"+str(timestamp).replace('/','.').replace(' ','h').replace(':','m')+".xlsx"


    run_Flux42_TC_EMP_02_1_1_Add_salary_decision(credentiale, fisierInput, sheetIn, fisierOutput, chromePath)
