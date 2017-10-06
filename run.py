# -*- encoding: utf8-*-

import os
import time
import requests
import re
import xlsxwriter
from lxml import etree

def parsing_page_content(content):
    # convert content to xml 
    xml_root = etree.XML(content)
    
    # parse component part 
    elements = etree.XPath('//component')(xml_root)[0]
    
    html = etree.HTML(elements.text)

    store_info_part = html.xpath('//table[@class="tiles"]//tr')


    store_info_cloumn = 0

    for value in store_info_part:
        field = value[0].text
        # get store name and store link info
        if field == "店名":
            store_field = field
            store_name = re.sub(r"\/|\s+", "", value[1][1].text)
            store_link = value[1][2].attrib['href'] if len(value[1]) > 3 else "None"
            print(store_field, store_name, store_link)
        # get location and GPS info
        elif field == "地址":
            address_field = field
            address_info = value[1][0][0].text
            address_latitude = "None"
            address_longitude = "None"
            if len(value[1][0]) > 2:
                geo_div = value[1][0][2]
                print(geo_div[0][0].text, geo_div[0][1].text)
                address_latitude = geo_div[0][0].text
                address_longitude = geo_div[0][1].text
            print(address_field, address_info, address_latitude, address_longitude)
        # get tel info
        elif field == "電話":
            tel_field = field
            tel_info = re.sub(r"\s+|\b-\b", "", value[1].text)
            print(tel_field, tel_info)
        # get delivery info
        elif field == "送達地區":
            delivery_field = field
            delivery_info = value[1].text
            print(delivery_field, delivery_info)
        # get order info
        elif field == "訂購說明":
            order_field = field
            order_info = value[1][1].text
            print(order_field, order_info)
        # get store service type
        elif field == "店家服務類型":
            service_field = field
            service_type = value[1].text
            print(service_field, service_type)
        # get create date info
        elif field == "最後修改日":
            update_field = "更新日期"
            update_time = value[1].text
            print(update_field, update_time)
            
    # Create an new Excel file and add a worksheet
    workbook = xlsxwriter.Workbook(os.getcwd() + '/output/' + store_name + '.xlsx')
    worksheet = workbook.add_worksheet()
    workformat = workbook.add_format({'text_wrap': True})
    category_format = workbook.add_format({'text_wrap': True})

    workformat.set_font_size(14)
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 60)
    worksheet.set_column('D:D', 40)
    worksheet.set_column('E:E', 40)

    worksheet.write(store_info_cloumn, 0, store_field, workformat)
    worksheet.write(store_info_cloumn, 1, store_name, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, "店家網址", workformat)
    worksheet.write(store_info_cloumn, 1, store_link if store_link != "None" else "", workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, address_field, workformat)
    worksheet.write(store_info_cloumn, 1, address_info, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, "位置", workformat)
    worksheet.write(store_info_cloumn, 1, address_latitude + ','  + address_longitude if address_latitude != "None" and address_longitude != "None" else "", workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, tel_field, workformat)
    worksheet.write(store_info_cloumn, 1, tel_info, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, delivery_field, workformat)
    worksheet.write(store_info_cloumn, 1, delivery_info, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, order_field, workformat)
    worksheet.write(store_info_cloumn, 1, order_info, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, service_field, workformat)
    worksheet.write(store_info_cloumn, 1, service_type, workformat)
    store_info_cloumn += 1
    worksheet.write(store_info_cloumn, 0, update_field, workformat)
    worksheet.write(store_info_cloumn, 1, update_time, workformat)

    menu_info_part = html.xpath('//table[@class="tiles"]')
   
    menu_info_cloumn = 0
   
    # get table of menu
    for table in menu_info_part[2:]:
        # get tr of table
        for tr in table:
            # It's category when value length is 1
            if len(tr) == 1:
                menu_category = tr[0][0].text
                print(tr[0][0].text)
                worksheet.write(menu_info_cloumn, 3, 'category', workformat)
                worksheet.write(menu_info_cloumn, 4, menu_category, workformat)
                menu_info_cloumn += 1
            # It's item of menu
            elif len(tr) == 2:
                # get name of item
                menu_item_name = tr[0][0].text
                print(tr[0][0].text)
                worksheet.write(menu_info_cloumn, 3, menu_item_name, workformat) 
                # get price of item
                menu_item_price = ""
                for value in tr[1]:
                    menu_item_price += value.text 
                print(menu_item_price)
                worksheet.write(menu_info_cloumn, 4, menu_item_price, workformat)
                menu_info_cloumn += 1
    workbook.close()

def crawler_store_list(stores_list):
    for value in stores_list:
        # get url element
        special_str = value.attrib['onclick']
        first_apostrophe = special_str.find("'") + 1
        second_apostrophe = special_str.find("'", first_apostrophe) 
        url = special_str[first_apostrophe:second_apostrophe]
        session.headers.update(store_headers)
        page = session.get("https://dinbendon.net" + url + "&random=0.26773047972369089")
        # unicode strings with encoding declaration are not supported. so that do not decode content
        parsing_page_content(page.content)

def get_next_page(url):
    session.headers.update(store_headers)
    res = session.get("https://dinbendon.net" + url + "&random=0.26773047972369089")

    # convert content to xml 
    xml_root = etree.XML(res.content)

    # parse component part 
    elements = etree.XPath('//component')(xml_root)[0]
    return etree.HTML(elements.text)

def get_next_page_link(html):
    # Get next page link
    next_link_element = html.xpath('//a[@id="navigation_panel_next"]')[0]
    next_link_str = next_link_element.attrib['onclick']
    first_index = next_link_str.find("'") + 1
    second_index = next_link_str.find("'", first_index)
    next_link = next_link_str[first_index: second_index]
    return next_link

main_page_headers = {
    "accept": "*/*",
    "cookie": "INDIVIDUAL_KEY=24032229-dcff-4d23-be76-c743d56459dd; ORIGINATOR_KEY=Jones; BUYER_DEFINES=NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL; MergeOrderItemShowComment=true; form.buyer=Jones; signIn.rememberMe=true; JSESSIONID=085D468309DE2F787AA6FD9A7C81AEAB; _ga=GA1.2.497461203.1461908205; _gid=GA1.2.387543421.1506912932; _gat=1",
    "refer": "https://dinbendon.net/do/login",
    "upgrade-insecure-requests": "1",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
    "accept-language": "zh-TW,zh;q=0.8,en-US;q=0.6,en;q=0.4",
    "accept-encoding": "gzip, deflate, sdch, br"
}
store_headers = {
    "accept": "*/*",
    "cookie": "INDIVIDUAL_KEY=24032229-dcff-4d23-be76-c743d56459dd; ORIGINATOR_KEY=Jones; BUYER_DEFINES=NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL__SEP__NULL; MergeOrderItemShowComment=true; form.buyer=Jones; signIn.rememberMe=true; JSESSIONID=085D468309DE2F787AA6FD9A7C81AEAB; _ga=GA1.2.497461203.1461908205; _gid=GA1.2.387543421.1506912932; _gat=1",
    "refer": "https://dinbendon.net/do/idine",
    "upgrade-insecure-requests": "1",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
    "accept-language": "zh-TW,zh;q=0.8,en-US;q=0.6,en;q=0.4",
    "accept-encoding": "gzip, deflate, sdch, br"
} 

STORE_LIST_URL = "https://dinbendon.net/do/idine"

session = requests.Session()

res = session.get(STORE_LIST_URL, headers=main_page_headers)

content = res.content.decode('utf8')

html = etree.HTML(res.content)

while html:
    # parse store url xpath
    store_list = html.xpath('//table[@class="ituneFrame grid"]//td[@class="cell"]//a')

    crawler_store_list(store_list)

    next_link = get_next_page_link(html)
    #print(next_link)
    html = get_next_page(next_link)
    #print(html)
    if not html:
        print('stop')
        break


"""
for value in result:
    # get url element
    special_str = value.attrib['onclick']
    first_apostrophe = special_str.find("'") + 1
    second_apostrophe = special_str.find("'", first_apostrophe) 
    url = special_str[first_apostrophe:second_apostrophe]
    session.headers.update(store_headers)
    page = session.get("https://dinbendon.net" + url + "&random=0.26773047972369089")
    # unicode strings with encoding declaration are not supported. so that do not decode content
    parsing_page_content(page.content)
""" 

