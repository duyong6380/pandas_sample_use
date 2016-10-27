#!/usr/sbin/env python
#-*- coding:utf-8 -*-
import pandas as pd
#import xlrd
from urllib2  import urlopen
import csv
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

def getExcelData():    
    FILENAME='2015-2016.xlsx'
    xls_file = pd.ExcelFile(FILENAME)
    last_list = ['2015','2016']
    for year in last_list:
        table = xls_file.parse(year)

        table = table[table[u'产品线.1'].isin(['AF',u'下一代防火墙'])]
        data = table[[u'客户行业',u'模块金额',u'订单号',u'产品数量',u'产品模块']]
        
        data_mouth = {'1':['01010000','03319999'],'2':['04010000','06309999'],\
                '3':['07010000','09309999'],'4':['10010000','12309999']}
        for key,mouth in data_mouth.items():
             max_num = int(year + mouth[1])
             min_num = int(year + mouth[0])
             mouth_data = data[(data[u'订单号'] < max_num) & (data[u'订单号'] > min_num)]
             mouth_data.groupby(data[u'客户行业']).sum().to_csv(year+'-'+key+'quarter_table.csv')

        try:
            data.groupby(data[u'客户行业']).sum().to_csv('duy_last_csvfile_trade_'+year+'.csv')
            tmp_data = data.groupby([data[u'客户行业'],data[u'产品模块']]).sum()
            print tmp_data
        except:
            print sys.exc_info()
        tmp_data.to_csv('duy_last_csvfile_module_'+year+'.csv')
    print 

def main():
    getExcelData()
    

if __name__ == '__main__':
    main()

