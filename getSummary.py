from selenium import webdriver
from bs4 import BeautifulSoup
sum_url="https://smmry.com/"
url="https://www.factcheck.org/2019/01/pelosi-didnt-spend-497-million-on-renovations/"

def get_article_sum(page):
    word=""
    for link in page.find_all('div',{'id':'sm_container_output'}):  
        word+=link.text
    return word
       
def summary(url):
    
    browser=webdriver.Chrome("C:\selenium\chromedriver.exe")
    browser.get(sum_url)
    
    
    url_element=browser.find_element_by_id('sm_force_url')
    url_element.send_keys(url)
    
    sub_element=browser.find_element_by_id('sm_submit')
    sub_element.submit()
            
    page=BeautifulSoup(browser.page_source,'html.parser')
    article_summary=get_article_sum(page)
    return article_summary
