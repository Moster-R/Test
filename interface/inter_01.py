#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @Time : 2020/9/4 23:05
# @Author : 墨
# @Version：V 0.1
# @File : inter_01.py
# @desc :
from interface.getexcl import ExclMth as ex
import os

class inter_01():
    dir_url = os.path.join(os.getcwd() + r'\test_case_api.xlsx')
    excl = ex(dir_url, 'login')
    data_Json = excl.ReadExcl()



    for data in data_Json:
        url = data[4]
        method = data[3]
        json = data[5]
        expected = eval(data[6])
        # print(url, method, json, expected)


        if json is not None:
            json = eval(json)


        res = ex.visa_json(

            url = url,
            method = method,
            json = json,
            headers={"Content-Type": "application/json", "X-Lemonban-Media-Type": "lemonban.v2"}
        )
        # print(res)
        res_msg = res["msg"]
        expected_msg = expected.get("msg")
        # print(res_msg,expected_msg)
        if res_msg == expected_msg:
            print("第{}条通过".format(data[0]))
            result = "PASS"

        else:
            print("第{}条不通过".format(data[0]))
            result = "FAILED"

        ex.SaveExcl(excl, data[0] + 1, result)




    
