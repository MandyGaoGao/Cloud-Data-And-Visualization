import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
import time
  
sum_url="https://smmry.com/"


def get_article_sum(page):
    word=""
    for link in page.find_all('div',{'id':'sm_container_output'}):  
        word+=link.text
    return word
       
def summary(url): 
    time.sleep(10)
    browser=webdriver.Chrome("C:\selenium\chromedriver.exe")
    browser.get(sum_url)
    url_element=browser.find_element_by_id('sm_force_url')
    url_element.send_keys(url)
    sub_element=browser.find_element_by_id('sm_submit')
    sub_element.submit()         
    page=BeautifulSoup(browser.page_source,'html.parser')
    article_summary=get_article_sum(page)
    return article_summary

def domain(url):
    #url="https://www.politifact.com/north-carolina/statements/2019/feb/08/jacob-wohl/trump-supporter-spreads-fake-news-about-nc-governo/"
    domain=""
    count=0
    for i in url:
        if count==3:
            break
        else:
            if i=="/":
                count+=1
            else:
                if count==2:
                    domain+=i
    return domain






#data=pd.read_excel(open('C:\Users\gaoyu\Desktop\DB.xlsx', 'rb'),sheetname='Sheet1') 
data = pd.read_excel(r'C:\Users\gaoyu\Desktop\DB.xlsx')
  
# creating a blank series #
data_new_summary = pd.Series([]) 
data_new_domain = pd.Series([]) 
data_new_truth = pd.Series([]) 
# running a for loop and asigning some values to series 
for i in range(10,90): 
    data_new_domain[i]=domain(data["url"][i])
    data_new_truth[i]="false"
    data_new_summary[i]=summary(data["url"][i])

data.insert(2, "tuthvalue", data_new_truth)                      
data.insert(4, "domain", data_new_domain)  
data.insert(5, "summary", data_new_summary)

data.to_excel("output2.xlsx")

    
def rate():
    rating="PolitiFact rating: Pants On Fire"
    rate=""
    count=0
    for i in rating:
        if i==":":
            count+=1
        else:
            if count==1:
                rate+=i
    print(rate)
