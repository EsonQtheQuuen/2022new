import re
import cProfile
import xlwt

from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.charts import Map

from bs4 import BeautifulSoup

import asyncio
from  pyppeteer import  launch
from pyppeteer import launcher

launcher.DEFAULT_ARGS.remove("--enable-automation")

async def pyppteer_fetchUrl(url):
    browser = await launch({'headless': False,'dumpio':True, 'autoClose':True})
    page = await browser.newPage()
    await page.goto(url)
    await asyncio.wait([page.waitForNavigation()])
    str = await page.content()
    await browser.close()
    return str

def fetchUrl(url):
    return asyncio.get_event_loop().run_until_complete(pyppteer_fetchUrl(url))
#获取html

def excel_built(data_dict,date):

    province = {
        "西藏": 0, "澳门": 0, "青海": 0, "台湾": 0, "香港": 0, "贵州": 0, "吉林": 0, "新疆": 0, "宁夏": 0, "内蒙古": 0,
        "甘肃": 0,
        "天津": 0, "山西": 0, "辽宁": 0, "黑龙江": 0, "海南": 0, "河北": 0, "陕西": 0, "云南": 0, "广西": 0, "福建": 0,
        "上海": 0,
        "北京": 0, "江苏": 0, "四川": 0, "山东": 0, "江西": 0, "重庆": 0, "安徽": 0, "湖南": 0, "河南": 0, "广东": 0,
        "浙江": 0, "湖北": 0
    }

    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet(date+'疫情数据')
    sheet.write(0, 0, '省市地区')
    sheet.write(0, 1, '新增确诊')
    sheet.write(0, 2, '新增无症状')
    count = 1

    for i in province.keys():
        sheet.write(count, 0, i)
        sheet.write(count, 1, data_dict['cfm_dict'][i])
        sheet.write(count, 2, data_dict['asy_dict'][i])
        count += 1
    workbook.save('Excel表.xls')

def chart_built(data_dict,date):
    chart = {
        Bar(init_opts=opts.InitOpts(width="1000px"))
        .add_xaxis(list(data_dict.keys()))
        .add_yaxis(date+"新增本土确诊",list(data_dict.values()))
        .set_global_opts(title_opts=opts.TitleOpts(title="今日疫情数据"))
        .render("疫情.html")
    }

def map_built(data_dict,date):
    province={
            "西藏":0,"澳门":0,"青海":0,"台湾":0,"香港":0,"贵州":0,"吉林":0,"新疆":0,"宁夏":0,"内蒙古":0,"甘肃":0,
                "天津":0,"山西":0,"辽宁":0,"黑龙江":0,"海南":0,"河北":0,"陕西":0,"云南":0,"广西":0,"福建":0,"上海":0,
                    "北京":0,"江苏":0,"四川":0,"山东":0,"江西":0,"重庆":0,"安徽":0,"湖南":0,"河南":0,"广东":0,"浙江":0,"湖北":0
              }
    for i in province.keys():
        if data_dict.get(i):
            province[i]=int(data_dict[i])
    china_map = {
        Map()
        .add(date+"今日新增", [list(z) for z in zip(province.keys(), province.values())], "china")
        .set_global_opts(title_opts=opts.TitleOpts(title="中国地图"),
                         visualmap_opts=opts.VisualMapOpts(min_=10, max_=100))
        .render("今日新增.html")
    }

def data_detail(str):#获取各个省市数据
    dict = {
        "西藏": 0, "澳门": 0, "青海": 0, "台湾": 0, "香港": 0, "贵州": 0, "吉林": 0, "新疆": 0, "宁夏": 0, "内蒙古": 0,
        "甘肃": 0,
        "天津": 0, "山西": 0, "辽宁": 0, "黑龙江": 0, "海南": 0, "河北": 0, "陕西": 0, "云南": 0, "广西": 0, "福建": 0,
        "上海": 0,
        "北京": 0, "江苏": 0, "四川": 0, "山东": 0, "江西": 0, "重庆": 0, "安徽": 0, "湖南": 0, "河南": 0, "广东": 0,
        "浙江": 0, "湖北": 0
    }
    provi= ["(西藏)(.*?)例", "(澳门)(.*?)例", "(青海)(.*?)例", "(台湾)(.*?)例", "(香港)(.*?)例", "(贵州)(.*?)例",
                "(吉林)(.*?)例", "(新疆)(.*?)例", "(宁夏)(.*?)例", "(内蒙古)(.*?)例",
                "(甘肃)(.*?)例", "(天津)(.*?)例", "(山西)(.*?)例", "(辽宁)(.*?)例", "(黑龙江)(.*?)例", "(海南)(.*?)例",
                "(河北)(.*?)例", "(陕西)(.*?)例", "(云南)(.*?)例", "(广西)(.*?)例",
                "(福建)(.*?)例", "(上海)(.*?)例", "(北京)(.*?)例", "(江苏)(.*?)例", "(四川)(.*?)例", "(山东)(.*?)例",
                "(江西)(.*?)例", "(重庆)(.*?)例", "(安徽)(.*?)例", "(湖南)(.*?)例",
                "(河南)(.*?)例", "(广东)(.*?)例", "(浙江)(.*?)例", "(湖北)(.*?)例"
                ]
    for item in provi:
        num=re.search(item,str)
        if num:
            dict[num.group(1)]=int(num.group(2))

    return dict

def data_get(mes):#处理文本中的数据

    yq_data={}

    cfm_data=re.search("31个省（自治区、直辖市）和新疆生产建设兵团报告新增确诊病例.*?例.*?本土(.*?)例（(.*?)）",mes)
    asy_data=re.search("31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者.*?例.*?本土(.*?)例（(.*?)）",mes)

    yq_data['cfm_dict']=data_detail(cfm_data.group(2))#返回数据字典
    yq_data['asy_dict']=data_detail(asy_data.group(2))

    yq_data['cfm']=re.search('(\d+)',cfm_data.group(1)).group(1)
    yq_data['asy']=re.search('(\d+)',asy_data.group(1)).group(1)

    return yq_data

def spider(url):#爬取当前通报所有文本
    html = fetchUrl(url)
    soup = BeautifulSoup(html, 'lxml')
    mes=''
    text = soup.find('div',attrs={'id':'xw_box'}).find_all('p')
    for i in text:
        mes+='\n'+i.text
    return mes

def special_area_data(mes1,mes2):#根据今日与昨日数据获取港澳台新增确诊数据
    dict={'香港':0,'澳门':0, '台湾':0}
    datalist=[]
    for i in range(1,4):
        data1 = re.search(
            '累计收到港澳台地区通报确诊病例.*?例。其中.*?香港特别行政区(.*?)例.*?澳门特别行政区(.*?)例.*?台湾地区(.*?)例',
            mes1).group(i)
        data2 = re.search(
            '累计收到港澳台地区通报确诊病例.*?例。其中.*?香港特别行政区(.*?)例.*?澳门特别行政区(.*?)例.*?台湾地区(.*?)例',
            mes2).group(i)
        datalist.append(int(data1)-int(data2))
    count=0
    for i in dict.keys():
        dict[i]=datalist[count]
        count+=1
    return dict

def main():
    url = 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml'  #需要请求的url
    html=fetchUrl(url)
    soup=BeautifulSoup(html,'lxml')
    page_url=soup.find_all('li')#爬取卫健委通报链接
    text=page_url[0].text
    date=re.search('(.*?)月(\d+)',text)
    date=date.group(1)+'月'+str(date.group(2))+'日'
    purl_today='http://www.nhc.gov.cn'+page_url[0].a['href']
    purl_yeasterday='http://www.nhc.gov.cn'+page_url[1].a['href']
    #爬取今日通报数据与昨日数据
    mes_today=spider(purl_today)
    mes_yeasterday=spider(purl_yeasterday)

    yq_dict=data_get(mes_today)
    chart_built(yq_dict['cfm_dict'], date)  # 今日新增确诊图表

    sp_data=special_area_data(mes_today,mes_yeasterday)
    for i in sp_data.keys():
            yq_dict['cfm_dict'][i]=sp_data[i]
    #处理今日新增数据

    map_built(yq_dict['cfm_dict'],date)#今日新增全国可视化
    excel_built(yq_dict, date)  # excel数据导入


if __name__ == '__main__':
    main()
    cProfile.run('main()')
