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


def switch2Window(driver, handle):
    for newhandle in driver.window_handles:
        if newhandle != handle:
            driver.switch_to.window(newhandle)


def switchBack(driver, handle):
    for newhandle in driver.window_handles:
        if newhandle != handle:
            driver.switch_to.window(newhandle)
            driver.close()

    driver.switch_to.window(handle)

def searchOneByName(driver, name):
    element = driver.find_element_by_id('searchkey')
    element.send_keys(name)
    element.send_keys(Keys.ENTER)

    elems = driver.find_elements_by_xpath('//div[contains(@class,"maininfo")]/a[1]')
    # 只取第一个
    detail_href = ''
    if len(elems) > 0:
        elems[0].click()
        return True
    else:
        return False

def queryOneDetail(driver):
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

    print("{}. {}".format(count, tmp_result))
    return tmp_result

count = 0
driver = None
handle = None
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
            handle = driver.current_window_handle

        if searchOneByName(driver, name=t_pha_name):
            switch2Window(driver, handle)
            queryOneDetail(driver)
            switchBack(driver, handle)
            driver.back()
        else:
            driver.back()

        count = count + 1

        if count % 20 == 19:
            driver.quit()
            driver = None
            handle = None

    except Exception as e:
        print(e)
        driver.quit()
        driver = None
        handle = None
        retry_count = retry_count + 1
        if retry_count > 3:
            print(t_pha_id)
            print(t_pha_name)
            print('something wrong, skip this one')
            count = count + 1
        continue

if driver is not None:
    # driver.quit()
    pass
