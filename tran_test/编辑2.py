# -*- coding: utf-8 -*-
"""
Created on Thu Oct 26 09:43:34 2017

@author: Administrator
"""

import urllib2 as request
from bs4 import BeautifulSoup
import  chardet
import re

if __name__ == '__main__':
    url = 'http://www.136book.com/huaqiangu/'
    head = {}
    #User-Agent头域的内容包含发出请求的用户信息。 
    head['User-Agent'] = 'Mozilla/5.0 (Linux; Android 4.1.1; Nexus 7 Build/JRO03D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166  Safari/535.19'
    req = request.Request(url, headers = head)
    response = request.urlopen(req)
    html = response.read()
    #print html
    soup = BeautifulSoup(html, 'lxml')
    #soup_texts = soup.find('div', id = 'book_detail', class_= 'box1').find_next('div')
    soup_texts = soup.find('div',id = 'book_detail', class_= 'box1').find_next('div')

    f = open('E:/huaqiangu_tmp.txt','w')
    for link in soup_texts.ol.children:
        if link != '\n':
            print link.text
            download_url = link.a.get('href')
            #print download_url
            download_req = request.Request(download_url, headers = head)
            download_response = request.urlopen(download_req)
            download_html = download_response.read()
            download_soup = BeautifulSoup(download_html, 'lxml')
            download_soup_texts = download_soup.find('div', id = 'content')
            
            # 抓取其中文本
            download_soup_texts = download_soup_texts.text.encode('utf-8')            
            
            #relast=r'(.*)document.write.*'
            #matches=re.findall(relast,download_soup_texts)
            #print matches
            
            begin=download_soup_texts.find('document.write')
            matches=''
            print begin
            matches=download_soup_texts[:begin]
            print matches
            
            # 写入章节标题
            str_tmp=link.text.encode('utf-8')
            f.write(str_tmp + '\n\n')
            # 写入章节内容
            f.write(matches)
            f.write('\n\n')
            
    f.close()
    
    f1 = open('E:/huaqiangu_log.txt','w')
    f1.write(str(link))
    f1.close()

