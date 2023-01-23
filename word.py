# -*- coding:utf-8 -*-
# @Time:2023/1/15 14:44
# @Author:Green
# @Function:此文件用于将题目导入、导出excel

import re,docx_to_strings,MD5,SM4,sql_operation

def main(doc:'试题1.docx'='试题1.docx') ->bool :

    password_txt = open('password.txt')                 #这是待会文件加密的密钥存放文件
    password = password_txt.read()

    sql_operation.login()                               #初始化数据库

    print('\n')
    docx_strings=docx_to_strings.main(doc)
    for row_index in range(len(docx_strings)):#row_index是当前遍历到的行数
        if '【单选题】' in docx_strings[row_index]:

            # 得到题目类型与题目得分
            temp0=re.findall('【(.*?)】.*?（(.*?)分）',docx_strings[row_index])


            question_class=temp0[0][0]
            question_score=temp0[0][1]

            question_body_hash=''
            question_body=''
            question_answer=''
            question_chapter=[]#此处用列表是因为一个问题可能属于多个章
            question_difficulty=0#难度一共分为1/3/5三个级别
            question_resolve=''#题目解析
            question_knowledge_points=''#题目知识点
            question_class='单选题'
            flag=1

            distinguish_a_A=''
            distinguish_sequence=''

            #接下来得到题目答案与题干（包括选项）
            while(flag!=0):
                if (('所属章：' or '所属章:') in docx_strings[row_index+flag]):
                    break
                if(re.findall('（([A-Z])）',docx_strings[row_index+flag])!=None and ('A、' in docx_strings[row_index+flag+1])):
                    question_answer=re.findall('（([A-Z])）',docx_strings[row_index+flag])[0]
                question_body+=docx_strings[row_index+flag]+'\n'
                flag+=1

            #接下来得到所属章、难度
            question_chapter_str=''#因为章节最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
            while(flag!=0):
                if((('难度：' or '难度:') in docx_strings[row_index+flag])==False):
                    question_chapter_str+=docx_strings[row_index+flag]
                else:
                    question_chapter_str=re.findall('所属章(.*?)$',question_chapter_str)[0]#把所属章几个字去掉
                    temp1 = re.findall('[、:：](.[^、]*?)$',question_chapter_str)#因为最后一个章节的字符后面没有'、'，所以要再匹配一次，进行交集运算
                    temp2 = re.findall('(.*?)、', question_chapter_str)
                    question_chapter = list(set(temp1).union(set(temp2)))#temp1和temp2做并集

                    #将章节与题干hash的表数据插入
                    question_body_hash = MD5.md5(question_body)
                    for chapter in question_chapter:
                        sql='insert into `question_bank`.`chapter` (`CHAPTER`,`QUESTION_HASH`) VALUES ("%s","%s")'
                        sql_operation.execute_sql_with_args(sql,(chapter,question_body_hash))

                    if '简单' in docx_strings[row_index+flag]:
                        question_difficulty=1
                    elif '一般' in docx_strings[row_index+flag]:
                        question_difficulty=3
                    else:
                        question_difficulty=5
                    flag+=1
                    break
                flag+=1

            # 接下来得到解析
            while (flag != 0):
                if (('知识点：' or '知识点:') in docx_strings[row_index + flag]):    # 因为解析最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                    question_resolve = re.findall('[:：](.*?)$', question_resolve)[0]  # 把解析几个字去掉
                    break
                else:
                    question_resolve += docx_strings[row_index + flag]
                flag += 1

            #接下来得到知识点
            while (flag != 0):
                if (docx_strings[row_index+flag]!=''):
                    question_knowledge_points += docx_strings[row_index + flag]    # 因为知识点最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                else:
                    question_knowledge_points = re.findall('[:：](.*?)$', question_knowledge_points)[0]  # 把解析几个字去掉
                    break
                flag += 1

            #对获取数据进行求hash、加密
            question_body_hash=MD5.md5(question_body)
            temp=SM4.SM4()
            question_body_code = temp.encrypt(password, question_body)
            question_body_code=str(question_body_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_resolve_code = temp.encrypt(password, question_resolve)
            question_resolve_code=str(question_resolve_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_knowledge_points_code = temp.encrypt(password, question_knowledge_points)
            question_knowledge_points_code=str(question_knowledge_points_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_answer_str_code = temp.encrypt(password, question_answer)
            question_answer_str_code=str(question_answer_str_code, encoding = "utf-8")

            #对有默认值要求的数据进行处理
            if distinguish_a_A=='':
                distinguish_a_A='否'
            if distinguish_sequence=='':
                distinguish_sequence='是'

            #本题所有数据上传数据库
            sql="INSERT INTO `question_bank`.`question` (`QUESTION_HASH`, `QUESTION_SCORE`, `QUESTION_ENCRAPT_BODY`, `DISTINGUISH_a_A`,`DISTINGUISH_ABC_CBA`,`QUESTION_RESOLVE`, `QUESTION_DIFFICULTY`, `QUESTION_KNOWLEDGE_POINT`,`QUESTION_CLASS`,`QUESTION_answer`) VALUES (\'{a}\', \'{b}\', \'{c}\', \'{x}\', \'{y}\', \'{d}\', \'{e}\', \'{f}\', \'{v}\', \'{w}\');".format(a=question_body_hash,b=question_score,c=question_body_code,d=question_resolve_code,e=question_difficulty,f=question_knowledge_points_code,x=distinguish_a_A,y=distinguish_sequence,v=question_class,w=question_answer_str_code)
            sql_operation.execute_sql(sql)

        elif '【多选题】' in docx_strings[row_index]:

            # 得到题目类型与题目得分
            temp0=re.findall('【(.*?)】.*?（(.*?)分）',docx_strings[row_index])


            question_class=temp0[0][0]
            question_score=temp0[0][1]

            question_body_hash=''
            question_body=''
            question_answer=''
            question_chapter=[]#此处用列表是因为一个问题可能属于多个章
            question_difficulty=0#难度一共分为1/3/5三个级别
            question_resolve=''#题目解析
            question_knowledge_points=''#题目知识点
            question_class='多选题'
            flag=1

            distinguish_a_A=''
            distinguish_sequence=''

            #接下来得到题目答案与题干（包括选项）
            while(flag!=0):
                if (('所属章：' or '所属章:') in docx_strings[row_index+flag]):
                    break
                if(re.findall('（([A-Z]*)）',docx_strings[row_index+flag])!=None and ('A、' in docx_strings[row_index+flag+1])):
                    question_answer=re.findall('（([A-Z]*)）',docx_strings[row_index+flag])[0]
                question_body+=docx_strings[row_index+flag]+'\n'
                flag+=1

            #接下来得到所属章、难度
            question_chapter_str=''#因为章节最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
            while(flag!=0):
                if((('难度：' or '难度:') in docx_strings[row_index+flag])==False):
                    question_chapter_str+=docx_strings[row_index+flag]
                else:
                    question_chapter_str=re.findall('所属章(.*?)$',question_chapter_str)[0]#把所属章几个字去掉
                    temp1 = re.findall('[、:：](.[^、]*?)$',question_chapter_str)#因为最后一个章节的字符后面没有'、'，所以要再匹配一次，进行交集运算
                    temp2 = re.findall('(.*?)、', question_chapter_str)
                    question_chapter = list(set(temp1).union(set(temp2)))#temp1和temp2做并集

                    #将章节与题干hash的表数据插入
                    question_body_hash = MD5.md5(question_body)
                    for chapter in question_chapter:
                        sql='insert into `question_bank`.`chapter` (`CHAPTER`,`QUESTION_HASH`) VALUES ("%s","%s")'
                        sql_operation.execute_sql_with_args(sql,(chapter,question_body_hash))

                    if '简单' in docx_strings[row_index+flag]:
                        question_difficulty=1
                    elif '一般' in docx_strings[row_index+flag]:
                        question_difficulty=3
                    else:
                        question_difficulty=5
                    flag+=1
                    break
                flag+=1

            # 接下来得到解析
            while (flag != 0):
                if (('知识点：' or '知识点:') in docx_strings[row_index + flag]):    # 因为解析最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                    question_resolve = re.findall('[:：](.*?)$', question_resolve)[0]  # 把解析几个字去掉
                    break
                else:
                    question_resolve += docx_strings[row_index + flag]
                flag += 1

            #接下来得到知识点
            while (flag != 0):
                if (docx_strings[row_index+flag]!=''):
                    question_knowledge_points += docx_strings[row_index + flag]    # 因为知识点最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                else:
                    question_knowledge_points = re.findall('[:：](.*?)$', question_knowledge_points)[0]  # 把解析几个字去掉
                    break
                flag += 1

            #对获取数据进行求hash、加密
            question_body_hash=MD5.md5(question_body)
            temp=SM4.SM4()
            question_body_code = temp.encrypt(password, question_body)
            question_body_code=str(question_body_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_resolve_code = temp.encrypt(password, question_resolve)
            question_resolve_code=str(question_resolve_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_knowledge_points_code = temp.encrypt(password, question_knowledge_points)
            question_knowledge_points_code=str(question_knowledge_points_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_answer_str_code = temp.encrypt(password, question_answer)
            question_answer_str_code=str(question_answer_str_code, encoding = "utf-8")

            #对有默认值要求的数据进行处理
            if distinguish_a_A=='':
                distinguish_a_A='否'
            if distinguish_sequence=='':
                distinguish_sequence='是'

            #本题所有数据上传数据库
            sql="INSERT INTO `question_bank`.`question` (`QUESTION_HASH`, `QUESTION_SCORE`, `QUESTION_ENCRAPT_BODY`, `DISTINGUISH_a_A`,`DISTINGUISH_ABC_CBA`,`QUESTION_RESOLVE`, `QUESTION_DIFFICULTY`, `QUESTION_KNOWLEDGE_POINT`,`QUESTION_CLASS`,`QUESTION_answer`) VALUES (\'{a}\', \'{b}\', \'{c}\', \'{x}\', \'{y}\', \'{d}\', \'{e}\', \'{f}\', \'{v}\', \'{w}\');".format(a=question_body_hash,b=question_score,c=question_body_code,d=question_resolve_code,e=question_difficulty,f=question_knowledge_points_code,x=distinguish_a_A,y=distinguish_sequence,v=question_class,w=question_answer_str_code)
            sql_operation.execute_sql(sql)

        elif '【填空题】' in docx_strings[row_index]:

            # 得到题目类型与题目得分
            temp0=re.findall('【(.*?)】.*?（每空(.*?)分）',docx_strings[row_index])


            question_class=temp0[0][0]
            question_score=temp0[0][1]

            question_body=''
            question_body_hash=''
            question_answer=[]
            distinguish_a_A=''# 是否区分答案英文大小写
            distinguish_sequence=''#是否区分填空题答案顺序
            question_chapter=[]#此处用列表是因为一个问题可能属于多个章
            question_difficulty=0#难度一共分为1/3/5三个级别
            question_resolve=''#题目解析
            question_knowledge_points=''#题目知识点
            question_answer_str='$'
            question_class='填空题'
            flag=1

            #接下来得到题目答案与题干
            while(flag!=0):
                if (('判分时答案内字母区分大小写：' or '判分时答案内字母区分大小写:') in docx_strings[row_index+flag]):
                    break
                question_body+=docx_strings[row_index+flag]+'\n'
                flag += 1
            question_answer.append(re.findall('【([\s\S]*?)】', question_body))

            #将答案变成字符串方便数据库直接存储，不用建立答案与题目对应的表
            for temp3 in question_answer[0]:
                question_answer_str+=temp3
            question_answer_str

            # 接下来得到题目判分时答案内字母区分大小写、判分时区分多个空的先后顺序
            distinguish_a_A=re.findall('[:：](.*?)$', docx_strings[row_index + flag])[0]
            flag += 1
            distinguish_sequence=re.findall('[:：](.*?)$', docx_strings[row_index + flag])[0]
            flag+=1

            #接下来得到所属章、难度
            question_chapter_str=''#因为章节最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
            while(flag!=0):
                if((('难度：' or '难度:') in docx_strings[row_index+flag])==False):
                    question_chapter_str+=docx_strings[row_index+flag]
                else:
                    question_chapter_str=re.findall('所属章(.*?)$',question_chapter_str)[0]#把所属章几个字去掉
                    temp1 = re.findall('[、:：](.[^、]*?)$',question_chapter_str)#因为最后一个章节的字符后面没有'、'，所以要再匹配一次，进行交集运算
                    temp2 = re.findall('(.*?)、', question_chapter_str)
                    question_chapter = list(set(temp1).union(set(temp2)))#temp1和temp2做并集

                    #将章节与题干hash的表数据插入
                    question_body_hash = MD5.md5(question_body)
                    for chapter in question_chapter:
                        sql='insert into `question_bank`.`chapter` (`CHAPTER`,`QUESTION_HASH`) VALUES ("%s","%s")'
                        sql_operation.execute_sql_with_args(sql,(chapter,question_body_hash))

                    if '简单' in docx_strings[row_index+flag]:
                        question_difficulty=1
                    elif '一般' in docx_strings[row_index+flag]:
                        question_difficulty=3
                    else:
                        question_difficulty=5
                    flag+=1
                    break
                flag+=1

            # 接下来得到解析
            while (flag != 0):
                if (('知识点：' or '知识点:') in docx_strings[row_index + flag]):    # 因为解析最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                    question_resolve = re.findall('[:：](.*?)$', question_resolve)[0]  # 把解析几个字去掉
                    break
                else:
                    question_resolve += docx_strings[row_index + flag]
                flag += 1

            #接下来得到知识点
            while (flag != 0):
                if (docx_strings[row_index+flag]!=''):
                    question_knowledge_points += docx_strings[row_index + flag]    # 因为知识点最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                else:
                    question_knowledge_points = re.findall('[:：](.*?)$', question_knowledge_points)[0]  # 把解析几个字去掉
                    break
                flag += 1

            #对获取数据进行求hash、加密
            question_body_hash=MD5.md5(question_body)
            temp=SM4.SM4()
            question_body_code = temp.encrypt(password, question_body)
            question_body_code=str(question_body_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_resolve_code = temp.encrypt(password, question_resolve)
            question_resolve_code=str(question_resolve_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_knowledge_points_code = temp.encrypt(password, question_knowledge_points)
            question_knowledge_points_code=str(question_knowledge_points_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_answer_str_code = temp.encrypt(password, question_answer_str)
            question_answer_str_code=str(question_answer_str_code, encoding = "utf-8")

            #对有默认值要求的数据进行处理
            if distinguish_a_A=='':
                distinguish_a_A='否'
            if distinguish_sequence=='':
                distinguish_sequence='是'

            #本题所有数据上传数据库
            sql="INSERT INTO `question_bank`.`question` (`QUESTION_HASH`, `QUESTION_SCORE`, `QUESTION_ENCRAPT_BODY`, `DISTINGUISH_a_A`,`DISTINGUISH_ABC_CBA`,`QUESTION_RESOLVE`, `QUESTION_DIFFICULTY`, `QUESTION_KNOWLEDGE_POINT`,`QUESTION_CLASS`,`QUESTION_answer`) VALUES (\'{a}\', \'{b}\', \'{c}\', \'{x}\', \'{y}\', \'{d}\', \'{e}\', \'{f}\', \'{v}\', \'{w}\');".format(a=question_body_hash,b=question_score,c=question_body_code,d=question_resolve_code,e=question_difficulty,f=question_knowledge_points_code,x=distinguish_a_A,y=distinguish_sequence,v=question_class,w=question_answer_str_code)
            sql_operation.execute_sql(sql)

        elif '【判断题】' in docx_strings[row_index]:

            # 得到题目类型与题目得分
            temp0=re.findall('【(.*?)】.*?（(.*?)分）',docx_strings[row_index])

            question_class=temp0[0][0]
            question_score=temp0[0][1]
            question_body_hash=''
            question_body=''
            question_answer=[]
            question_chapter=[]#此处用列表是因为一个问题可能属于多个章
            question_difficulty=0#难度一共分为1/3/5三个级别
            question_resolve=''#题目解析
            question_knowledge_points=''#题目知识点
            question_class='判断题'
            flag=1

            distinguish_a_A=''
            distinguish_sequence=''

            #接下来得到题目答案与题干
            while(flag!=0):
                if (('所属章：' or '所属章:') in docx_strings[row_index+flag]):
                    break
                question_body+=docx_strings[row_index+flag]+'\n'
                flag += 1
            question_answer=re.findall('（(.)）', question_body)

            #接下来得到所属章、难度
            question_chapter_str=''#因为章节最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
            while(flag!=0):
                if((('难度：' or '难度:') in docx_strings[row_index+flag])==False):
                    question_chapter_str+=docx_strings[row_index+flag]
                else:
                    question_chapter_str=re.findall('所属章(.*?)$',question_chapter_str)[0]#把所属章几个字去掉
                    temp1 = re.findall('[、:：](.[^、]*?)$',question_chapter_str)#因为最后一个章节的字符后面没有'、'，所以要再匹配一次，进行交集运算
                    temp2 = re.findall('(.*?)、', question_chapter_str)
                    question_chapter = list(set(temp1).union(set(temp2)))#temp1和temp2做并集

                    #将章节与题干hash的表数据插入
                    question_body_hash = MD5.md5(question_body)
                    for chapter in question_chapter:
                        sql='insert into `question_bank`.`chapter` (`CHAPTER`,`QUESTION_HASH`) VALUES ("%s","%s")'
                        sql_operation.execute_sql_with_args(sql,(chapter,question_body_hash))

                    if '简单' in docx_strings[row_index+flag]:
                        question_difficulty=1
                    elif '一般' in docx_strings[row_index+flag]:
                        question_difficulty=3
                    else:
                        question_difficulty=5
                    flag+=1
                    break
                flag+=1

            # 接下来得到解析
            while (flag != 0):
                if (('知识点：' or '知识点:') in docx_strings[row_index + flag]):    # 因为解析最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                    question_resolve = re.findall('[:：](.*?)$', question_resolve)[0]  # 把解析几个字去掉
                    break
                else:
                    question_resolve += docx_strings[row_index + flag]
                flag += 1

            #接下来得到知识点
            while (flag != 0):
                if (docx_strings[row_index+flag]!=''):
                    question_knowledge_points += docx_strings[row_index + flag]    # 因为知识点最多一百个字，如果字数很多时可能章节会转行，正则表达式不能正确识别，所以先把几行进行检测拼接
                else:
                    question_knowledge_points = re.findall('[:：](.*?)$', question_knowledge_points)[0]  # 把解析几个字去掉
                    break
                flag += 1

            #对获取数据进行求hash、加密
            question_body_hash=MD5.md5(question_body)
            temp=SM4.SM4()
            question_body_code = temp.encrypt(password, question_body)
            question_body_code=str(question_body_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_resolve_code = temp.encrypt(password, question_resolve)
            question_resolve_code=str(question_resolve_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_knowledge_points_code = temp.encrypt(password, question_knowledge_points)
            question_knowledge_points_code=str(question_knowledge_points_code, encoding = "utf-8")
            temp=SM4.SM4()
            question_answer_str_code = temp.encrypt(password, question_answer)
            question_answer_str_code=str(question_answer_str_code, encoding = "utf-8")

            #对有默认值要求的数据进行处理
            if distinguish_a_A=='':
                distinguish_a_A='否'
            if distinguish_sequence=='':
                distinguish_sequence='是'

            #本题所有数据上传数据库
            sql="INSERT INTO `question_bank`.`question` (`QUESTION_HASH`, `QUESTION_SCORE`, `QUESTION_ENCRAPT_BODY`, `DISTINGUISH_a_A`,`DISTINGUISH_ABC_CBA`,`QUESTION_RESOLVE`, `QUESTION_DIFFICULTY`, `QUESTION_KNOWLEDGE_POINT`,`QUESTION_CLASS`,`QUESTION_answer`) VALUES (\'{a}\', \'{b}\', \'{c}\', \'{x}\', \'{y}\', \'{d}\', \'{e}\', \'{f}\', \'{v}\', \'{w}\');".format(a=question_body_hash,b=question_score,c=question_body_code,d=question_resolve_code,e=question_difficulty,f=question_knowledge_points_code,x=distinguish_a_A,y=distinguish_sequence,v=question_class,w=question_answer_str_code)
            sql_operation.execute_sql(sql)

    print('word.main finished')
    return True

if __name__=='__main__':
    result=main()