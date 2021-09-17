from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.Excel.Application import Application
import time
import re
from bs4 import BeautifulSoup
import os
import shutil
import glob
import pandas as pd
class ITDASHBOARD():
    def __init__(self,url,download_dire):
        self.curr=os.getcwd()
        self.browser = Selenium()
        self.url=url
        self.download_dir=download_dire

        self.download_dir1=self.curr
        self.browser.set_download_directory(self.curr)
        self.browser.open_available_browser(self.url)

    def get_web_page(self):
        self.btn_xpath='//*[@id="node-23"]/div/div/div/div/div/div/div/a'

        

        self.browser.click_link(self.btn_xpath)
        time.sleep(10)
        self.page_source=self.browser.get_source

        self.doc=BeautifulSoup(self.page_source(),'html.parser')

        self.soup1=self.doc.find_all('div',{'id':'agency-tiles-2-container'})
        return self.soup1

    def get_agencies_data(self):

        self.name_and_money=(self.soup1[0].find_all('span'))

        self.name=[]
        self.money=[]
        for i in self.soup1[0].find_all('span',{'class':'h4 w200'}):
            self.name.append(i.text)
        for i in self.soup1[0].find_all('span',{'class':'h1 w900'}):
            self.money.append(i.text)
            self.data1=zip(self.name,self.money)
            self.data=dict(self.data1)
            self.df=pd.DataFrame(list(zip(self.name,self.money)))
            self.df.columns=['Agency Name','Agency Spendings']

            self.df.to_excel('Agencies.xlsx')
    def get_individual_investments(self):
        
            
        self.browser.click_link('view')
        time.sleep(10)
        self.browser.select_from_list_by_label('investments-table-object_length','All')
        time.sleep(15)
        self.page_source2=self.browser.get_source
        self.doc2=BeautifulSoup(self.page_source2(),'html.parser')
        self.soup2=self.doc2.find('div',{'id':'investments-table-container'})
        self.UII=[]
        for i in self.soup2.find_all('td',{'class':'left sorting_2'}):
            self.UII.append(i.text)

        self.Bureau=[]
        for i in self.soup2.find_all('td',{'class':'left select-filter'}):
            self.Bureau.append(i.text)

        self.Inv_title=[]
        for i in self.soup2.find_all('td',{'class':'left'}):
            self.Inv_title.append(i.text)
        self.Total_Spend=[]
        for i in self.soup2.find_all('td',{'class':'right'}):
            self.Total_Spend.append(i.text)
                         

                         
        self.CIO_Rat=[]
        for i in self.soup2.find_all('td',{'class':'center'}):
            self.CIO_Rat.append(i.text)

        self.bureau=[]
        self.cio=[]
                         
        self.type1=[]
        self.No_of_proj=[]
        for count, i in enumerate(self.Bureau):
            if count % 2 != 1:
                self.bureau.append(i)
            else:
                self.type1.append(i)                         

        for count, i in enumerate(self.CIO_Rat):
            if count % 2 != 1:
                self.cio.append(i)
            else:
                self.No_of_proj.append(i)

        for i in self.UII:
            for j in self.Inv_title:
                if i==j:
                    self.Inv_title.remove(j)
        for i in self.bureau:
            for j in self.Inv_title:
                if i==j:
                    self.Inv_title.remove(j)

        for i in self.type1:
            for j in self.Inv_title:
                if i==j:
                    self.Inv_title.remove(j)

        self.data23=pd.DataFrame()
        self.data23['UII']=self.UII
        self.data23['Bureau']=self.bureau
        self.data23['Investment Title']=self.Inv_title
        self.data23['Total FY2021 Spending ($M)']=self.Total_Spend
        self.data23['CIO Rating']=self.cio
        self.data23['# of Projects']=self.No_of_proj

        self.data23.to_excel('Individual Investments.xlsx')
    def get_pdfs(self):

        self.page_source3=self.browser.get_source

        self.a=self.browser.get_webelements('xpath://a')
        self.l1=[]
        for i in self.a:
            self.l1.append(self.browser.get_text(i))
        for i in self.l1:
            if '005' not in i:
                self.l1.remove(i)
        self.l2=[]
        for i in self.l1:
            if '005-' in i:
                self.l2.append(i)
        self.links=[]
        for a in self.soup2.find_all('a',href=True):
            self.links.append(a['href'])

        self.ll1=[]
        self.len_links=len(self.links)
        self.half_links=self.len_links//2
        for i in range(3):
            self.ll1.append(self.links[i])
        for i in self.ll1:
            self.up=self.url+i
            self.browser.go_to(self.up)
            time.sleep(10)
            self.browser.click_link('Download Business Case PDF')
            time.sleep(5)
            self.browser.go_back()
            self.browser.go_back()
            time.sleep(10)
            self.browser.select_from_list_by_label('investments-table-object_length','All')
            time.sleep(2)
    def save_into_pdfs(self):
        for i in glob.glob(self.curr+'/*.pdf'):
            shutil.move(i,self.curr+'/output')
        for i in glob.glob(self.curr+'/*.xlsx'):
            shutil.move(i,self.curr+'/output')
        self.browser.close_all_browsers()
download_dir="/output"
url = "https://itdashboard.gov/"
make=ITDASHBOARD(url,download_dir)
make.get_web_page()
make.get_agencies_data()
make.get_individual_investments()
make.get_pdfs()
make.save_into_pdfs()


