#!/usr/local/bin/python3
# coding=UTF-8

import openpyxl
import re
import datetime
import sqlite3
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.action_chains import ActionChains


g_test_batch = 2

# 1. query data base on the db index
cx = sqlite3.connect('./result/operations.db')
cx.execute("create table if not exists qcc_firm_idx (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, qccdetailhref TEXT)")
cx.execute("create table if not exists qcc_firm_detail (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, \
    qccname TEXT, qccstatus TEXT, qcctags TEXT, qccused TEXT, qccnewtags TEXT, qcctel TEXT, qccoffweb TEXT, qccemail TEXT, qccadd TEXT)")
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

    # company status
    status_elems = driver.find_elements_by_xpath('//div[contains(@class,"title")]/div/span/span[1]')
    tmp_result['qccstatus'] = ('无' if len(status_elems) == 0 else status_elems[0].text)

    # company tags
    tmp_result['qcctags'] = []
    tmp_result['qccused'] = ""
    company_tags = driver.find_elements_by_xpath('//div[contains(@class,"tags")]/span[contains(@class, "text-primary")]')
    for tag in company_tags:
        if tag.text == '曾用名':
            hover = ActionChains(driver).move_to_element(tag)
            hover.perform()   #悬停
            company_used_names = driver.find_elements_by_xpath('//div[contains(@class,"tags")]/span[contains(@class, "text-primary")]/div')
            tmp_result['qccused'] = ','.join(list(map(lambda x: x.text, company_used_names))).replace('\n', '')
        else:
            tmp_result['qcctags'].append(tag.text)

    tmp_result['qcctags'] = ','.join(tmp_result['qcctags'])

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

def result2DB(id, name, d):
    sql = "insert into qcc_firm_detail (phaid, phaname, qccname , qccstatus , qcctags , qccused , qccnewtags , qcctel , " \
          "qccoffweb, qccemail, qccadd) VALUES ('{}', '{}', '{}', '{}', '{}', '{}','{}', '{}', '{}', '{}', '{}')" \
            .format(id, name, d['qccname'], d['qccstatus'], d['qcctags'], d['qccused'], d['qccnewtags'], \
                    d['qcctel'], d['qccoffweb'], d['qccemail'], d['qccadd'])
    print(sql)
    cur.execute(sql)
    cx.commit()

def error2DB(id, name):
    sql = "insert into qcc_firm_detail (phaid, phaname, qccname , qccstatus , qcctags , qccused , qccnewtags , qcctel , " \
          "qccoffweb, qccemail, qccadd) VALUES ('{}', '{}', '{}', '{}', '{}', '{}','{}', '{}', '{}', '{}', '{}')" \
        .format(id, name, '', '', '', '', '', '', '', '', '')
    print(sql)
    cur.execute(sql)
    cx.commit()


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
            result = queryOneDetail(driver)
            switchBack(driver, handle)
            result2DB(t_pha_id, t_pha_name, result)
            driver.back()
        else:
            driver.back()

        count = count + 1

        if count % 20 == 19:
            driver.quit()
            driver = None
            handle = None

    except Exception as e:
        driver.quit()
        driver = None
        handle = None
        retry_count = retry_count + 1
        if retry_count > 3:
            error2DB(t_pha_id, t_pha_name)
            count = count + 1
        continue

if driver is not None:
    driver.quit()
