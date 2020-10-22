# !/usr/bin/env python
# _*_ coding:utf-8 _*_
__author__ = 'Gomjie'

import xlrd


class getexcel:
    def __init__(self, filepath, xlsname, sheetname):
        self.path = filepath
        self.xlsname = xlsname
        self.sheetname = sheetname

    def get_excel(self):
        """
        加载execl文件数据
        :param path: 文件路径
        :return:返回数据
        """
        try:
            dataresult = []  # 保存从excel表中读取出来的值，每一行为一个list，dataresult中保存了所有行的内容
            result = []  # 是由dict组成的list，是将dataresult中的内容全部转成字典组成的list：result
            datapath = self.path + '\\' + self.xlsname
            xls1 = xlrd.open_workbook(datapath)
            table = xls1.sheet_by_name(self.sheetname)
            for i in range(0, table.nrows):
                dataresult.append(table.row_values(i))
            # 将list转化成dict
            for i in range(1, len(dataresult)):
                temp = dict(zip(dataresult[0], dataresult[i]))
                result.append(temp)
            return result
        except Exception as msg:
            print("异常消息-> {0}".format(msg))

    def alldata(self):
        """
        读取execl文件数据
        :return: 返回数据
        """
        data = self.get_excel()
        return data

    def caselen(self):
        """
        testcase字典长度
        :return: 字典长度大小
        """
        data = self.alldata()
        length = len(data)
        return length

    def checklen(self):
        """
        check字典长度
        :return: 字典长度大小
        """
        data = self.alldata()
        length = len(data)
        return length

    def get_elementinfo(self, i):
        """
        获取testcase项的element_info元素
        :param i: 位置序列号
        :return: 返回element_info元素数据
        """
        data = self.alldata()
        return data[i]['element_info']

    def get_find_type(self, i):
        """
        获取testcase项的find_type元素数据
        :param i: 位置序列号
        :return: 返回find_type元素数据
        """
        data = self.alldata()
        return data[i]['find_type']

    def get_operate_type(self, i):
        """
        获取testcase项的operate_type元素数据
        :param i: 位置序列号
        :return: 返回operate_type元素数据
        """
        data = self.alldata()
        return data[i]['operate_type']

    def get_CheckElementinfo(self, i):
        """
        获取check项的element_info元素
        :param i: 位置序列号
        :return: 返回element_info元素数据
        """
        data = self.alldata()
        return data[i]['element_info']

    def get_CheckFindType(self, i):
        """
        获取check项的element_info元素
        :param i: 位置序列号
        :return: 返回find_type元素数据
        """
        data = self.alldata()
        return data[i]['find_type']

    def get_CheckOperate_type(self, i):
        data = self.alldata()
        return data[i]['operate_type']
