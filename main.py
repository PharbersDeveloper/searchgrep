#!/usr/local/bin/python3
# coding=UTF-8

import openpyxl
import re
import datetime
import sqlite3
from selenium import webdriver
from selenium.webdriver import Keys


# 1. query data base on the db index
cx = sqlite3.connect('./result/operations.db')
cx.execute("create table if not exists qcc_firm_idx (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, qccdetailhref TEXT)")
cx.execute("create table if not exists qcc_firm_detail (id INTEGER PRIMARY KEY autoincrement, phaid TEXT, phaname TEXT, \
    qcctp TEXT, qcctag TEXT, qccboss TEXT, qccphone TEXT, qcccode TEXT, qccoffweb TEXT, qccemail TEXT, qccadd TEXT)")
cur = cx.cursor()



# driver = webdriver.Chrome(executable_path='lib/chromedriver')
# driver.implicitly_wait(10)     # seconds
# driver.get('http://www.baidu.com')
# driver.get('http://www.qcc.com')
