# -*- coding:utf-8 -*-
# @Time:2023/1/18 7:58
# @Author:Green
# @Function:此函数用于将图片、公式转换为str，并将docx全部转换为strings

import docx
import os, re, hashlib, sys, shutil, sql_operation

def get_pictures(word_path:'输入存图片的文档路径'='./试题1.docx', result_path:'输入图片保存目录路径'='./pictures')->bool:
    try:
        if not os.path.exists(result_path):
            os.mkdir(result_path)
        shutil.rmtree(result_path)  # 清空缓存图片的文件夹，以防图片数量超过文档中的（以防上次导入题目的图片剩余）
        doc = docx.Document(word_path)
        dict_rel = doc.part._rels
        for rel in dict_rel:
            rel = dict_rel[rel]
            if "image" in rel.target_ref:
                if not os.path.exists(result_path):
                    os.makedirs(result_path)
                else:
                    pass
                img_name = re.findall("/(.*)", rel.target_ref)[0]
                word_name = os.path.splitext(word_path)[0]
                if os.sep in word_name:
                    new_name = word_name.split('\\')[-1]
                else:
                    new_name = word_name.split('/')[-1]
                img_name = f'{new_name}-'+'.'+f'{img_name}'
                with open(f'{result_path}/{img_name}', "wb") as f:
                    f.write(rel.target_part.blob)
        print('docx_to_strings.get_pictures finished')
        return True
    except Exception as e:
        print(e)

def pic_to_hash(pic_dir:'str:输入图片的存储目录路径'='./pictures')->list:

    sql_operation.login()#初始化数据库

    filenames = os.listdir(pic_dir)                 #遍历目录文件夹
    pics_hashs=[]
    for file_name in filenames:
        # 读取文件
        file_name=pic_dir+'/'+file_name
        with open(file_name, 'rb') as fp:
            data = fp.read()                        #data是图片的数据流
        # 使用 md5 算法
        file_md5 = hashlib.md5(data).hexdigest()
        pics_hashs.append(file_md5)
        sql="INSERT INTO `QUESTION_BANK`.`IMAGE` (`PIC_HASH`,`IMAGEDATA`) VALUES(\'{A}\',%s);".format(A=file_md5)
        fp.close()
        sql_operation.execute_sql_with_args(sql,data)
    return pics_hashs

def docx_str(doc_path='./试题1.docx')->list:
    pics_hashs_list=pic_to_hash()#此处使用默认图片存储文件夹
    doc = docx.Document(doc_path)
    part = doc.part
    element = part.element
    xml = element.xml
    strxml = ''#用于记录当前行docx的所有文本信息，主要包括（文本字符、图片的hash、线性公式）
    for str1 in xml:
        strxml += str(str1)
    strxml = strxml.replace('\n', '')
    temp1=re.findall('<w:p(.*?)/w:p>',strxml)#temp1是111个段落标签内的xml标签集合
    the_text_around_equation=[]#通过<w:p(.*?)/w:p>把每行解析出来
    for temp2 in temp1:#temp2是单个段落标签内的xml标签集合
        pic_index=0#记录当前遇到的图片次数
        if 'docPr' in temp2:#检测是否有图片标签的特征
            tempstr = ''#清空每行的字符串缓存变量
            temp01=re.findall('(.*?)<pic:pic',temp2)#temp01是本行中第一个带有文本（线性公式、文本）的标签到图片标签之间的所有xml信息
            temp02=re.findall('/pic:pic>(.*?)',temp2)#temp02是本行中图片标签到第一个带有文本（线性公式、文本）的标签之间的所有xml信息
            if(temp01!=None):
                strings_before_pic= re.findall('<[^d]:t>(.*?)</[^d]:t>', temp01[0])
            if(temp02!=None):
                strings_after_pic= re.findall('<[^d]:t>(.*?)</[^d]:t>', temp02[0])
            if(temp01!=None):
                for i in strings_before_pic:
                    tempstr+=i
            tempstr+='${' + pics_hashs_list[pic_index] + '}$'
            if(temp02!=None):
                for j in strings_after_pic:
                    tempstr+=j
            pic_index+=1
            the_text_around_equation.append(tempstr)#这一行结束
        else:
            temp3=re.findall('<[^d]:t>(.*?)</[^d]:t>',temp2)#通过<[^d]:t>这样的文本标识把公式、文本解析出来
            tempstr=''
            for temp4 in temp3:
                tempstr+=temp4
            the_text_around_equation.append(tempstr)

    #下面是模块功能检查
    # for i in the_text_around_equation:
    #     print(i)
    print("docx_to_strings.docx_str finished")
    return the_text_around_equation

def main(doc_path:'试题1.docx'='./试题1.docx')->list:#返回值是docx完全转str后的列表，每行为一个元素
    get_pictures()
    strings=docx_str(doc_path)
    print("docx_to_strings.main finished")
    return strings

if __name__ == '__main__':
    main('./试题1.docx')



