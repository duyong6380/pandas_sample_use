#!/usr/sbin/env python
#-*- coding:utf-8 -*-
import pandas as pd
#import xlrd
from urllib2  import urlopen
import csv
import sys

csvfile = file('csvtes2.csv', 'wb')
writer = csv.writer(csvfile)
writer.writerow(['ID','版本号','订单号','最终用户','最终用户联系人','最终用户联系电话','行业','区域','销售人员名称','订货单位','产品型号'])

def writercsv(table):
    global writer
    global csvfile
    try:
       writer.writerows(table)
    except:
        print sys.exc_info()
def parse_line_data(linedata):
    ''''''
 #   print linedata
    data = [linedata['devId'],linedata['pdVersion'],linedata['tid'],linedata['finalCustomer'],\
            linedata['finalUser'],linedata['finalTel'],linedata['tradeName'],\
            linedata['areaName'],linedata['crmUserName'],linedata['orderUnit'],\
            linedata['pdName']]
    return data

def getExcelData():    
    FILENAME='2015-2016.xlsx'
    xls_file = pd.ExcelFile(FILENAME)
    table = xls_file.parse('2016')
    '''
    data = {}
    data['ID'] = table[u'订单号']
    print data
    data['trade'] = table[u'客户行业']
    data['product'] = table[u'产品模块']
    data['sum'] = table[u'模块金额']
    print data
    '''
    print len(table)
    return table

def process_column(tabledata):
    data = {}
    last_data = {}
    sumth = {}
    for index in range(len(tabledata)):
        tmp_data  = tabledata.ix[index]
        data['ID'] = tmp_data[u'订单号']
        data['trade'] = tmp_data[u'客户行业']
        data['product'] = tmp_data[u'产品模块']
        data['sum'] = tmp_data[u'模块金额']
        last_data[index] = data

    length = len(last_data)
    for index in range(length):
        if 
    print len(last_data)
    
def main():
    process_column(getExcelData())
    

if __name__ == '__main__':
    main()

