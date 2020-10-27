#!/usr/bin/env python
# -*- coding: cp1252  -*-

import selenium
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import psycopg2,psycopg2.extras
import wx,sys
from datetime import date
import xlwt
reload(sys)
sys.setdefaultencoding("cp1252")


class Traitement():
    def __init__(self):
        app=wx.App(0)

        driver = webdriver.Firefox() # Get local session of firefox
        driver.implicitly_wait(5)
        driver.get("http://www.salon-aps.com/site/FR/Exposants__Partenaires,C5268,I5104,SType-DIRECT.htm?KM_Session=47702c0a79f777c3bd985f51b021df72") # Load page
        chgpage=driver.find_element_by_xpath("//div[@class='chgpage']")
        a=chgpage.find_elements_by_tag_name("a")
        t_a_text=[" "]
        tzvalueline=[]
        for t in range(len(a)):
            t_a_text.append(a[t].text)

        for i in range(len(t_a_text)):
            try:
                if i==len(t_a_text)-1:
                    break
                if i>0:
                    print ""
                    print "-------- page: -------- ", t_a_text[i]
                    print ""
                    chgpage=driver.find_element_by_xpath("//div[@class='chgpage']")
                    cli=chgpage.find_element_by_link_text(t_a_text[i])
                    cli.click()

                #driver.find_element_by_css_selector("b.soc").click()
                div=driver.find_element_by_id("listtoprint")

                societes=div.find_elements_by_css_selector("b.soc")

                j=0
                t_text=[]
                for txt in societes:
                    t_text.append(txt.text)

                for x in range(len(t_text)):
                    div=driver.find_element_by_id("listtoprint")
                    soc=div.find_element_by_link_text(t_text[x])
                    soc.click()
                    sRs=""
                    sAdr=""
                    sPays=""
                    sVille=""
                    sCP=""
                    sTel=""
                    sFax=""
                    sWebsite=""
                    sEmail=""
                    sNom=""
                    sPrenom=""
                    sFonction=""
                    sCoords=""
                    cpville=""
                    tvalue=[]
                    infosoc=driver.find_element_by_xpath("//div[@class='desc']/div[@class='infosoc']")
                    sRs=str(infosoc.text.encode("cp1252")).strip()
                    print ""
                    print "Rs: ", sRs

                    coords=driver.find_element_by_xpath("//div[@class='desc']/div[@class='coords']")
                    #nbre_br=coords.find_elements_by_tag_name("br")
                    #print "len(nbre_br): ",len(nbre_br)

                    #sCoords=str(coords.text.encode("cp1252")).strip()
                    sCoords=coords.text.encode("cp1252")

                    svar1 = str(sCoords).split("\n")
                    n=0
                    #vide=False
                    iCp=0
                    p=0
                    for y in range(len(svar1)):
                        svar2=str(svar1[y].encode("cp1252")).strip()

                        if svar2.find("Tél :")!=-1:
                            p=y
                            sTel= svar2[len("Tél :"):]
                            sTel=str(sTel).strip()
                        if svar2.find("Fax :")!=-1:
                            sFax = svar2[len("Fax :"):]
                            sFax=str(sFax).strip()
                        if svar2.find("Web :")!=-1:
                            sWebsite = svar2[len("Web :"):]
                            sWebsite=str(sWebsite).strip()
                        if svar2.find("E-mail :")!=-1:
                            sEmail = svar2[len("E-mail :"):]
                            sEmail=str(sEmail).strip()


                    if p<>0:
                        for w in range(len(svar1)):
                            svar2=str(svar1[w].encode("cp1252")).strip()
                            if w==p-1:
                                sPays=str(svar2).strip()
                            if w==p-2:
                                cpville = str(svar2).split(" ")
                                for c in range(len(cpville)):
                                    if c==0:
                                        sCP=cpville[c]
                                        sCP=str(sCP).strip()
                                    else:
                                        sVille= str(sVille+" "+cpville[c]).strip()

                            if w>=0 and w<p-2:
                                sAdr= sAdr+" "+svar2
                                sAdr=sAdr.replace("\n","")
                                sAdr= str(sAdr).strip()

                    print "Adresse: ", sAdr
                    print "scp: ", sCP
                    print "sVille: ", sVille
                    print "sPays: ", sPays
                    print "sTel: ", sTel
                    print "sFax: ", sFax
                    print "sWebsite: ", sWebsite
                    print "sEmail: ", sEmail
                    tvalue.append(sRs)
                    tvalue.append(sAdr)
                    tvalue.append(sPays)
                    tvalue.append(sVille)
                    tvalue.append(sCP)
                    tvalue.append(sTel)
                    tvalue.append(sFax)
                    tvalue.append(sWebsite)
                    tvalue.append(sEmail)
                    tvalue.append("")
                    tvalue.append("")
                    tvalue.append("")
                    tzvalueline.append(tvalue)
                    driver.back()

                    #if j==2:
                    #    break
                    j=j+1

                #if i==2:
                #    break

            except NoSuchElementException, e:
                print e

        #for txt in t_text:
        #    try:
        #        lien=driver.find_element_by_xpath("//b[@class='soc'][@text="+txt+"]")
        #        print lien
        #    except NoSuchElementException, e:
        #        print e


        tzdata=[]
        tztitre=[]
        tzentete=["Raison sociale","Adresse","Pays","Ville","Code Postale","Téléphone","Fax","Website","E-mail","Nom Personne Contact","Prénom Personne Contact","Fonction"]
        tzfeuille=["salon aps"]
        tzdata.append(tzvalueline)
        tztitre.append(tzentete)
        dateNow = date.today()
        if self.doxls (tztitre,tzdata,"c:\\users\\salon aps_"+str(dateNow)+".xls",tzfeuille)==False:
            return False


        message="Traitement terminé dans "+"c:\\users\\salon aps_"+str(dateNow)+".xls"
        message=message.encode('cp1252')
        msg = wx.MessageDialog ( None, message, caption="VIVETIC",style=wx.OK , pos=wx.DefaultPosition )
        msg.ShowModal()


    def doxls(self,tzTitre,tzDatas,path,tzfeuille):
        wb = xlwt.Workbook(encoding='cp1252',style_compression=0)
        nbfeille=len(tzfeuille)
        F=0
        while (F<nbfeille):
            ws = wb.add_sheet(tzfeuille[F])
            I=0
            while (I<len(tzTitre[F])):
                ws.write(0, I, tzTitre[F][I])
                I=I+1

            K=0
            while(K<len(tzDatas[F])):
                L=0
                while (L <len(tzDatas[F][K])):

                    try:
                        ws.write(K+1, L, (tzDatas[F][K][L]))
                    except Exception as inst:
                        msgs =  'type ERREUR:'+str(type(inst))+'\n'     # the exception instance
                        msgs+=  'CONTENU:'+str(inst)+'\n'           # __str__ allows args to printed directly
                        self.traitement.errorDlg('erreur',msgs)
                        return False
                    L=L+1
                K=K+1

            F=F+1
        wb.save(path)
        return True



    def dlgprogress(self,parent,stitle='',smessage='',max=100,):
        dlg = wx.ProgressDialog(stitle,smessage,maximum = max,parent=parent,style = wx.PD_AUTO_HIDE|wx.PD_CAN_ABORT| wx.PD_APP_MODAL| wx.PD_ESTIMATED_TIME| wx.PD_REMAINING_TIME)
        return dlg

if __name__ == "__main__":
     Traitement()










#elem.send_keys("seleniumhq" + Keys.RETURN)
#time.sleep(0.2) # Let the page load, will be added to the API
#try:
#    browser.find_element_by_xpath("//a[contains(@href,'http://seleniumhq.org')]")
#except NoSuchElementException:
#assert 0, "can't find seleniumhq"
#browser.close()
