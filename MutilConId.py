#!/usr/sbin/env python
#-*- coding:utf-8 -*-
import pandas as pd
#import xlrd
from urllib2  import urlopen
import csv
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import json

csvfile = file('csvtes2.csv', 'wb')
writer = csv.writer(csvfile)
writer.writerow(['ID','版本号','订单号','最终用户','最终用户联系人','最终用户联系电话','行业','区域','销售人员名称','订货单位','产品型号','下游渠道'])

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
            linedata['pdName'],linedata['agentName']]
    return data

def getIdByUrl(serial_num):
    table_data = []
    url = 'http://200.200.0.56/sf/index.php/order/orderTasks/getSerialList?key=2016json&gwid='+serial_num
    lastdata = urlopen(url).read()
  #  print lastdata
    data = json.loads(lastdata)
    table_data = [parse_line_data(line) for line in data ]
    return  writercsv(table_data)

def getExcelData():    
    FILENAME='AF111.xlsx'
    xls_file = pd.ExcelFile(FILENAME)
    table = xls_file.parse('Sheet1')
    return table['gataway_ID']

def getTableDataByID(tabledata):
    count = 0
    serial_num = ""
    for data in tabledata:
        data = str(data).strip()
        if len(data) != 8:
            print '=============data:[',data,']======='
            print len(data)
            continue

        if count % 500 == 0 and count != 0:
           # print serial_num
            getIdByUrl(serial_num)
            serial_num = ""

        if count % 500 != 0:
            serial_num += ','
        serial_num +=data
        count += 1


def main():
    getTableDataByID(getExcelData())

if __name__ == '__main__':
    main()

exit( 0)    

