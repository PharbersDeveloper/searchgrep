#!/usr/local/bin/python3
# coding=UTF-8

import openpyxl
import re
import datetime
import sqlite3
from selenium import webdriver
from selenium.webdriver import Keys


g_test_batch = 100

# 1. query data base on the db index
cx = sqlite3.connect('./result/operations.db')
cx.execute("create table if not exists qcc_firm_idx (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, qccdetailhref TEXT)")
cx.execute("create table if not exists qcc_firm_detail (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, \
    qcctp TEXT, qcctag TEXT, qccboss TEXT, qccphone TEXT, qcccode TEXT, qccoffweb TEXT, qccemail TEXT, qccadd TEXT)")
cur = cx.cursor()

cur.execute("select count(*) from qcc_firm_detail")
g_skip_rows = cur.fetchall()[0][0]
print(g_skip_rows)

cur.execute("select * from qcc_firm_idx limit " + str(g_test_batch) + " offset " + str(g_skip_rows))
indeices = cur.fetchall()

def searchOneByName(driver):
    pass

def queryOneDetail(driver):
    driver.get(t_qcc_url)

    tmp_result = {}
    # company title
    title_elems = driver.find_elements_by_xpath('//div[contains(@class,"title")]/div/span/h1[1]')
    tmp_result['qccname'] = title_elems = title_elems[0].text
    # company tags
    company_tags = driver.find_elements_by_xpath('//div[contains(@class,"tags")]/span[contains(@class, "text-primary")]')
    tmp_result['qcctags'] = ','.join(list(map(lambda x: x.text, company_tags)))
    # company district
    company_address_tags = driver.find_elements_by_xpath('//div[contains(@class,"newtags")]')
    tmp_result['qccnewtags'] = ','.join(list(map(lambda x: x.text, company_address_tags)))

    # company contact-info
    company_info = driver.find_elements_by_xpath('//div[contains(@class,"contact-info")]/div[@class="rline"]/span[contains(@class, "f")]')
    for item in company_info:
        tmp = item.text.replace('\n', '')
        if '电话' in tmp:
            tmp_result['qcctel'] = tmp
        elif '邮箱' in tmp:
            tmp_result['qccemail'] = tmp
        elif '官网' in tmp:
            tmp_result['qccoffweb'] = tmp
        elif '地址' in tmp:
            tmp = tmp.replace('附近企业', '')
            tmp_result['qccadd'] = tmp
        else:
            print('loss message')
            print(tmp)

    # TODO: to database
    print("{}. {}".format(count, tmp_result))

count = 0
driver = None
retry_count = 0
while count < len(indeices):
    try:
        idx = indeices[count]
        t_pha_id = idx[1]
        t_pha_name = idx[2]
        t_qcc_url = idx[3]
        print(t_qcc_url)

        if driver is None:
            driver = webdriver.Chrome(executable_path='lib/chromedriver')
            driver.implicitly_wait(10)     # seconds
            driver.get("https://www.qcc.com")

        queryOneDetail(driver)
        count = count + 1

    except Exception as e:
        driver.quit()
        driver = None
        continue

if driver is not None:
    # driver.quit()
    pass
