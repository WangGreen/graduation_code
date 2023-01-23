# -*- coding:utf-8 -*-
# @Time:2023/1/18 15:20
# @Author:Green
# @Function:此文件用于进行对字符串的加解密，需要输入密码我已保存在password.txt文件中

#-*-coding: utf-8-*-

from gmssl.sm4 import CryptSM4, SM4_ENCRYPT, SM4_DECRYPT
import binascii
import base64

class SM4:
    """
    国密sm4加解密
    """

    def __init__(self):
        self.crypt_sm4 = CryptSM4()

    def str_to_hexStr(self, hex_str):
        """
        字符串转hex
        :param hex_str: 字符串
        :return: hex
        """
        hex_data = hex_str.encode('gbk')
        str_bin = binascii.unhexlify(hex_data)
        return str_bin.decode('gbk')

    def encrypt(self, encrypt_key, value):
        """
        国密sm4加密
        :param encrypt_key: sm4加密key
        :param value: 待加密的字符串
        :return: sm4加密后的hex值
        """
        crypt_sm4 = self.crypt_sm4
        crypt_sm4.set_key(encrypt_key.encode(), SM4_ENCRYPT)
        date_str = str(value)
        encrypt_value = crypt_sm4.crypt_ecb(date_str.encode())  # bytes类型
        #return encrypt_value.hex()
        return base64.b64encode(encrypt_value)

    def decrypt(self, decrypt_key, encrypt_value):
        """
        国密sm4解密
        :param decrypt_key:sm4加密key
        :param encrypt_value: 待解密的hex值
        :return: 原字符串
        """
        crypt_sm4 = self.crypt_sm4
        crypt_sm4.set_key(decrypt_key.encode(), SM4_DECRYPT)
        #decrypt_value = crypt_sm4.crypt_ecb(bytes.fromhex(encrypt_value))  # bytes类型
        decrypt_value = crypt_sm4.crypt_ecb(base64.b64decode(encrypt_value))
        return self.str_to_hexStr(decrypt_value.hex())

if __name__ == '__main__':
    password_txt =open('password.txt')
    password = password_txt.read()
    str_data = "HELLO"
    SM4 = SM4()
    print("待加密内容：", str_data)
    print("密钥为：",password)
    encoding = SM4.encrypt(password, str_data)
    print("国密sm4加密后的结果：", encoding)
    print("国密sm4解密后的结果：", SM4.decrypt(password, encoding))

