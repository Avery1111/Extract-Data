import requests
import urllib.request
import json
import time
from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO
from io import open
import string
import re
from bs4 import BeautifulSoup as bs
import datetime


FindUrl_doc = re.compile(r'var\s*_iframeUrl\s*=\s*(.*?;)')
FindUrl_key = re.compile(r'WOPISrc=(.*?);')

headers = {
    
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36'

}



def get_Date():

    datetoday = datetime.datetime.today().strftime('%Y-%m-%d')
    end_time = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    
    start_year = int(time.strftime('%Y',time.localtime(time.time()))) - 1
    month_day = time.strftime('%m-%d',time.localtime(time.time()))
    start_time = '{}-{}'.format(start_year,month_day)


    return {'datetoday':datetoday,'lastyear_today':start_time,'seDate':''+start_time+'~'+datetoday+''}




def readPDF(pdfFile):       
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    process_pdf(rsrcmgr, device, pdfFile)
    device.close()
    
    content = retstr.getvalue()
    retstr.close()
    return content


def get_doc_pdf(date,ID,suffix):

    url_base = "https://view.officeapps.live.com/op/view.aspx?src=http%3A%2F%2Fstatic.cninfo.com.cn%2Ffinalpage%2F{}%2F{}.{}".format(date,ID,suffix)
    # print(url_base)
    response_base = requests.get(url_base,headers=headers)
    html_base = response_base.text
    soup = bs(html_base,'html.parser')
    script = str(soup.script)
    
    url_doc_unicode = re.findall(FindUrl_doc,script)[0]
    url_doc = url_doc_unicode.encode('utf-8').decode('unicode_escape')
    url_key = re.findall(FindUrl_key,url_doc)[0]
    url_pdf = 'https://psg3-word-view.officeapps.live.com/wv/WordViewer/request.pdf?WOPIsrc={}&type=downloadpdf'.format(url_key)
    req = urllib.request.Request(url= url_pdf,headers=headers)
    response = urllib.request.urlopen(req)
    time.sleep(0.5)
    Text = readPDF(response)
    response.close()

    return Text



def get_FinalPage(url,seDate):

    query = {
        
        'pageNum': 1,
        'pageSize': 30,
        'column': 'szse',
        'tabName': 'relation',
        'plate': '',
        'stock': '',
        'searchkey': '',
        'secid': '',
        'category': '',
        'trade': '',
        'seDate': seDate,
        'sortName': '',
        'sortType': '',
        'isHLtitle': 'true'
    
    }
    
    
    response = requests.post(url,data=query,headers=headers)
    rawdata = response.json()
    # print(rawdata)
    finalpage = int(rawdata['totalpages'])
    
    return finalpage





def get_data(url,page,seDate):

    lst = []
    query = {
        
        'pageNum': page,
        'pageSize': 30,
        'column': 'szse',
        'tabName': 'relation',
        'plate': '',
        'stock': '',
        'searchkey': '',
        'secid': '',
        'category': '',
        'trade': '',
        'seDate': seDate,
        'sortName': '',
        'sortType': '',
        'isHLtitle': 'true'
    
    }
        
        
    response = requests.post(url,data=query,headers=headers)
    try:
        rawdata = response.json()
    except:
        return lst

    DataList = rawdata['announcements']
    # print(DataList)
    

    for d,item in enumerate(DataList):
        # print(item)
        adjuncturl = item['adjunctUrl']
        ID = item['announcementId']
        pubtime_stamp = int(item['announcementTime']/1000)
        
        pubtime_array = time.localtime(pubtime_stamp)
        pubtime = time.strftime('%Y-%m-%d',pubtime_array)
    
        if item['adjunctType'] == 'PDF':
        
            pdf_url = 'http://www.cninfo.com.cn/new/announcement/download?bulletinId={}&announceTime={}'.format(ID,pubtime)
            # print(item['announcementTitle'])
            req = urllib.request.Request(pdf_url,headers=headers)
            response_pdf = urllib.request.urlopen(req)
            time.sleep(0.5)
        

            try:
                Text = readPDF(response_pdf)
            except KeyError:
                Text = ''
                pass
            except:
                pass

            response_pdf.close()

        elif item['adjunctType'] == 'DOC':
            try:
                try:
                    Text = get_doc_pdf(date=pubtime,ID=ID,suffix='doc')
                except IndexError:
                    Text = get_doc_pdf(date=pubtime,ID=ID,suffix='DOC')
            except:
                Text = ''
        try:
            dic = {'code':item['secCode'],'name':item['secName'],'标题公告':item['announcementTitle'],'时间':pubtime,'文本':Text}
        except UnboundLocalError:
            dic = {'code':item['secCode'],'name':item['secName'],'标题公告':item['announcementTitle'],'时间':pubtime,'文本':''}
        # print(dic)
        lst.append(dic)
        print(f"{d+1} item's got")
        
    return lst

 


def save_json(result,path):

    with open(path,'w') as jf:
        json.dump(result,jf)
        jf.close()









if __name__ == '__main__':

    url = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'

    finalpage = get_FinalPage(url,get_Date()['seDate'])
    result = []
    for page in range(1,finalpage+1):
        print(f'\ngetting notice in page {page}/{finalpage}')
        data = get_data(url=url,page=page,seDate=get_Date()['seDate'])
        result.append(data)


        # if page>1:
        #     break

    result = [i for item in result for i in item]
    save_json(result=result,path='.\\data/Result.json')
    print('DATA SAVED')



#读取json文件
    # with open('.\\data/Result.json','r') as jr:
    #     output = json.load(jr)
    #     jr.close()
    
    # print(output)