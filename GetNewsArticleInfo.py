from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup
import re
from openpyxl import Workbook
import pandas as pd


url1 =[ 'https://mediabiasfactcheck.com/2019/01/24/the-latest-fact-checks-from-the-international-fact-checking-network-ifcn-1-24-2019/']
#'https://mediabiasfactcheck.com/2019/01/31/the-latest-fact-checks-from-the-international-fact-checking-network-ifcn-1-31-2019/']
#url ='https://mediabiasfactcheck.com/2019/02/09/the-latest-fact-checks-from-the-international-fact-checking-network-ifcn-2-9-2019/'
web='https://mediabiasfactcheck.com/category/fact-check-2/'
def get_urls(web):
    links=[]
    uClient=uReq(web)
    '''download the web'''
    page_html=uClient.read()
    uClient.close()
    '''page parse'''
    page_soup=soup(page_html,'html.parser')
    main=page_soup.find('div',id="mh-loop")   
    for link in main.find_all('a'):
        url=link.get('href')
        if 'the-latest-fact-checks-from-the-international-fact-checking-network-ifcn-'in url:
            links.append(url)    
    return links

def catchinfo(url):
    lst3=[]
    for my_url in url:
        uClient=uReq(my_url)
        '''download the web'''
        page_html=uClient.read()
        uClient.close()
        '''page parse'''
        page_soup=soup(page_html,'html.parser') 
        articletext=page_soup.text
        lst=page_soup.find_all(string=re.compile("Claim"))   
        a_position=articletext.index(lst[0])
        
        start_at = -1
        while True:            
            try:                
                b_position=articletext.index(lst[-1],start_at+1)
            except ValueError:
                break
            else:
                start_at=b_position
      
        descrip=str(articletext[int(a_position):int(b_position)])
        lst2=descrip.split('\n')
        for i in lst2:
            if i!='':
                lst3.append(i)

    return lst3

def catchArticleUrl(url):
    
    links=[]
    for i in url:
        uClient=uReq(i)
        '''download the web'''
        page_html=uClient.read()
        uClient.close()
        '''page parse'''
        page_soup=soup(page_html,'html.parser')
        mains=page_soup.find_all('div',{"class":"fc-viewarticle"})   
        
        for main in mains[:-1]:
            for link in main.find_all('a',{'class':'fc-clickable fc-viewarticle'},
                                      {'title':'View article in a new window'}):
                url=link.get('href')
                links.append(url)    
    return links

    
    
    
def savetoEXCEL(lst3,lst4):
    ls1=[]
    ls2=[]
    ls3=[]
    for i in lst3:
        if lst3.index(i)%3==0:
            ls1.append(i)
        if lst3.index(i)%3==1:
            ls2.append(i)
        if lst3.index(i)%3==2:
            ls3.append(i)    
    wb=Workbook()
    ws=wb.active
    ws.append(ls1)
    ws.append(ls2)
    ws.append(ls3)
    ws.append(lst4)
    wb.save('data.xlsx')

def trans(file):
    df=pd.read_excel(file)
    df=df.transpose()
    df.to_excel("data2.xlsx")
    
def test(web):    
    for i in range(2):
        try:
            url=get_urls(web)
            lst4=catchArticleUrl(url)
            lst3=catchinfo(url)
            print('case'+str(i))    
            url=get_urls(web)[:i]
            #print(url[-1])      
            #lst4=catchArticleUrl(url)
            #lst3=catchinfo(url)
            #print(lst3)
            print(len(lst4))
            print(len(lst3))
            print(len(lst3)%len(lst4))
        except:
            break

#url=get_urls(web)[:24]
#print(url[-1])      
#lst4=catchArticleUrl(url)
#lst3=catchinfo(url)
#print(lst3)
#print(len(lst4))
#print(len(lst3))
#savetoEXCEL(lst3,lst4)
file="data.xlsx"
trans(file)
