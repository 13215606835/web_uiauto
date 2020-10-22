#!/usr/bin/env python
# _*_ coding:utf-8 _*_
__author__ = 'Guomengjie'

import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(__file__)))
import unittest, ddt, xlrd
from config import setting_execl
from public.models import myunit, screenshot
from public.page_obj.loginPage_execl import login
from public.models.log import Log

try:
    xls2 = xlrd.open_workbook(setting_execl.TEST_DATA_EXECL + '/' + 'login_data.xlsx')
    table = xls2.sheet_by_name('Sheet1')
    dataresult = []
    for i in range(0, table.nrows):
        dataresult.append(table.row_values(i))
    # 将list转化成dict
    result = []
    for i in range(1, len(dataresult)):
        temp = dict(zip(dataresult[0], dataresult[i]))
        result.append(temp)
    testData = result
except FileNotFoundError as file:
    log = Log()
    log.error("文件不存在：{0}".format(file))


@ddt.ddt
class Demo_UI(myunit.MyTest):
    """登录测试"""

    def user_login_verify(self, phone, password):
        """
        用户登录
        :param phone: 手机号
        :param password: 密码
        :return:
        """
        login(self.driver).user_login(phone, password)

    def exit_login_check(self):
        """
        退出登录
        :return:
        """
        login(self.driver).login_exit()

    @ddt.data(*testData)
    def test_login(self, dataexecl):
        """
        登录测试
        :param dataexecl: 加载login_data登录测试数据
        :return:
        """
        log = Log()
        log.info("当前执行测试用例ID-> {0} ; 测试点-> {1}".format(dataexecl['id'], dataexecl['phone'], dataexecl['password']))
        # 调用登录方法
        self.user_login_verify(dataexecl['phone'], dataexecl['password'])
        po = login(self.driver)
        if dataexecl['screenshot'] == 'phone_pawd_success':
            log.info("检查点-> {0}".format(po.user_login_success_hint()))
            self.assertEqual(po.user_login_success_hint(), dataexecl['check'],
                             "成功登录，返回实际结果是->: {0}".format(po.user_login_success_hint()))
            log.info("成功登录，返回实际结果是->: {0}".format(po.user_login_success_hint()))
            screenshot.insert_img(self.driver, dataexecl['screenshot'] + '.jpg')
            log.info("-----> 开始执行退出流程操作")
            self.exit_login_check()
            po_exit = login(self.driver)
            log.info("检查点-> 找到{0}元素,表示退出成功！".format(po_exit.exit_login_success_hint()))
            self.assertEqual(po_exit.exit_login_success_hint(), '登录',
                             "退出登录，返回实际结果是->: {0}".format(po_exit.exit_login_success_hint()))
            log.info("退出登录，返回实际结果是->: {0}".format(po_exit.exit_login_success_hint()))
        else:
            log.info("检查点-> {0}".format(po.phone_pawd_error_hint()))
            self.assertEqual(po.phone_pawd_error_hint(), dataexecl['check'],
                             "异常登录，返回实际结果是->: {0}".format(po.phone_pawd_error_hint()))
            log.info("异常登录，返回实际结果是->: {0}".format(po.phone_pawd_error_hint()))
            screenshot.insert_img(self.driver, dataexecl['screenshot'] + '.jpg')


if __name__ == '__main__':
    unittest.main()
