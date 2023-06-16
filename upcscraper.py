from lxml import html
import requests
import pandas as pd
from biip.upc import Upc
from tor_proxy import tor_proxy

port = tor_proxy()

http_proxy  = f"socks5h://127.0.0.1:{port}"
https_proxy = f"socks5h://127.0.0.1:{port}"

proxies = { 
              "http"  : http_proxy, 
              "https" : https_proxy, 
            }


def PadUPC(upc,requiredlen):
    upc=str(upc)
    return ("0"*(requiredlen-len(upc)))+upc

def Process_UPC(upc):
    upc=str(upc)
    length=len(upc)
    if length<6:
        upc=PadUPC(upc,6)
    elif length>8 and length<12:
        upc=PadUPC(upc,12)
    return("https://www.iga.net/en/product/00000_"+PadUPC(Upc.parse(upc).payload,18))






urllist=[]
upclist=[]
pointslist=[]
igaurllist=[]

def ScrapeAllMetroItems(times):


    for k in range(times):
        prefix="https://www.metro.ca/en/online-grocery/search"
        pointpath='/div[1]/a/div[2]/div/span[2]//text()'
        linkpath='/div[1]/a/@href'
        shortxpath='//*[@id="content-temp"]/div/div[3]/div[3]/div/div[{}]'
        pages='//*[@id="content-temp"]/div/div[3]/div[4]/div/div/a[6]//text()'
        print("scraped products "+str(k)+" times")
        
        while True:
            try:
                page = requests.get(prefix,proxies=proxies)
                break
            except Exception:
                pass
        print(pages)
        tree = html.fromstring(page.content)
        pages=int(tree.xpath(pages)[0])
        pagecounter=0
        for j in range(1,pages+1):
            if j==1:
                page = requests.get(prefix+"?sortOrder=relevance&filter=%3Arelevance%3Adeal%3Ametro%26moi+Promotions&__cf_chl_tk=e9XmKyohKo4X16bnBkCX78Vv0i14vgDOI30XRo2M64I-1685554713-0-gaNycGzNFSU")
            else:
                page=requests.get(prefix+"-page-"+str(j)+"?sortOrder=relevance&filter=%3Arelevance%3Adeal%3Ametro%26moi+Promotions&__cf_chl_tk=e9XmKyohKo4X16bnBkCX78Vv0i14vgDOI30XRo2M64I-1685554713-0-gaNycGzNFSU")
            tree = html.fromstring(page.content)
            counter=0
            for i in range(100000):
                thispath=shortxpath.format(i)
                points = tree.xpath(thispath+pointpath)
                url=tree.xpath(thispath+linkpath)
                #print(url)
                if url==[]:
                    counter+=1
                if counter>=10:
                    pagecounter+=1
                    break
                if points!=[]:
                    url=prefix+url[0]
                    urllist.append(url)
                    upc=str(url.split("/")[-1:][0])
                    upclist.append(upc)
                    pointslist.append(points[0])
                    igaurllist.append(Process_UPC(upc))
                else:
                    continue
                pagecounter=0
                print(points,url)

ScrapeAllMetroItems(100)
df=pd.DataFrame(data={"UPC":upclist,"Metro URL":urllist,"IGA URL":igaurllist,"Points":pointslist})
df=df.drop_duplicates()
df.to_csv("C:\\Users\\desjardav\\Downloads\\metropoints.csv")

# print(page.status_code)

# points = tree.xpath(ex2)
# print(points)
# /html/body/div[1]/div[9]/div[2]/div[2]/div[2]/div/div[1]/div/div/div[3]/div[3]/div/div[1]
# /html/body/div[1]/div[9]/div[2]/div[2]/div[2]/div/div[1]/div/div/div[3]/div[3]/div/div[2]
