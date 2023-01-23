# -*- coding:utf-8 -*-
# @Time:2023/1/22 17:17
# @Author:Green
# @Function:此文件用于对输入的字符串求MD5的hash值并返回

import hashlib

def md5(string:"待求md5的字符串")->str:
    result = hashlib.md5(string.encode('utf-8')).hexdigest()
    return result

if __name__ == '__main__':
    a=md5('hello')