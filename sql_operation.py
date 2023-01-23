# -*- coding:utf-8 -*-
# @Time:2023/1/14 22:40
# @Author:Green
# @Function:此文件用于进行对数据库的增删改查

import pymysql,time

#此处登录数据库，初始化环境，创建表、属性
def login() ->bool:   #建立数据表，单选题（题目hash为主键、分值、题干密文、难度、解析、知识点），图片（图片hash为主键、图片流数据）
    # ，章节（章节hash为主键、加密密文），章节对应单选题表（表id为主键、章节hash、单选题hash）
    sql=[]
    db = pymysql.connect(host='localhost',
                     user='root',
                     password='Mydb2021+',
                     )
    cursor = db.cursor()
    sql.append( 'CREATE DATABASE IF NOT EXISTS `question_bank` DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;')

    sql.append('USE `question_bank`;')#进入数据库

    # 创建章节的表
    sql.append(
        'CREATE TABLE IF NOT EXISTS `CHAPTER`(CHAPTER VARCHAR(30) , QUESTION_HASH VARCHAR(255) ,'
        'PRIMARY KEY(QUESTION_HASH,CHAPTER)) DEFAULT charset=utf8;')

    #创建图片的表
    sql.append('create table if not  exists IMAGE(PIC_HASH varchar(255) ,IMAGEDATA mediumblob,PRIMARY KEY(PIC_HASH)) DEFAULT charset=utf8; ')

    # 创建题目的表,因为题干中图片的hash值比较长，所以题干密文长度设置为600
    sql.append('CREATE TABLE IF NOT EXISTS `QUESTION`(QUESTION_HASH VARCHAR(255),QUESTION_SCORE INT,QUESTION_CLASS VARCHAR(16) ,' \
               'QUESTION_ENCRAPT_BODY VARCHAR(600),DISTINGUISH_a_A VARCHAR(6) DEFAULT \'否\',DISTINGUISH_ABC_CBA VARCHAR(6)  DEFAULT \'是\','\
               'QUESTION_RESOLVE VARCHAR(255),QUESTION_DIFFICULTY INT,QUESTION_KNOWLEDGE_POINT VARCHAR(255),QUESTION_ANSWER VARCHAR(100),'
               'PRIMARY KEY(QUESTION_HASH)) DEFAULT charset=utf8;')

    try:
        for temp in sql:
            cursor.execute(temp)
    except Exception as e:
        log = open('mysql_err_log.txt', 'a')
        e=str(e)
        e='\n'+e+'\n'
        log.write(e)
        localtime = time.localtime(time.time())
        log.write(str(localtime))


def execute_sql(sql):
    db = pymysql.connect(host='localhost',
                     user='root',
                     password='Mydb2021+',
                    database='question_bank'
                     )
    cursor = db.cursor()
    try:
        cursor.execute(sql)
        db.commit()
    except Exception as e:
        log = open('mysql_err_log.txt', 'a')
        e = str(e)
        e = '\n' + e + '\n'
        log.write(e)
        localtime = time.localtime(time.time())
        print(e)
        log.write(str(localtime))

def execute_sql_with_args(sql,x):#图片存储和普通sql语句有点参数区别，所以重新写一个,x是图片的比特流
    db = pymysql.connect(host='localhost',
                     user='root',
                     password='Mydb2021+',
                    database='question_bank'
                     )
    cursor = db.cursor()
    try:
        cursor.execute(sql,(x))
        db.commit()
    except Exception as e:
        log = open('mysql_err_log.txt', 'a')
        e = str(e)
        e = '\n' + e + '\n'
        log.write(e)
        localtime = time.localtime(time.time())
        print(e)
        log.write(str(localtime))

def main():
    login()
    pass

if __name__ == "__main__":
    main()