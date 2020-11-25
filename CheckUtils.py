# -*- coding: utf-8 -*-

# 文件io
import os

from ConfigUtils import Config


# 字符串校验
class StrUtil():
    
    # 判空
    def isBlack(self, c_str):
        if (not isinstance(c_str, str)):
            raise RuntimeError("传入参数不是字符串！")
        if (c_str is None or c_str.strip() is ""):
            return True
        else:
            return False

    # 判不为空
    def isNotBlack(self, c_str):
        return not self.isBlack(c_str)


class FileUtil():
    __strUtils = StrUtil()

    # 文件是否存在
    def isExist(self, file_name):
        # 文件名是否为空
        if (self.__strUtils.isBlack(file_name)): 
            raise RuntimeError("文件名不能为空！")
        # 文件是否存在
        if os.path.exists(file_name):
            return True
        else:
            print("文件不存在！")
            return False

def test():
    # 字符串测试
    print("----------StrUtil测试------------")
    s  = StrUtil()
    print("s.isNotBlack('  ')测试结果：{}".format(s.isNotBlack("  ")))
    try:
        print("s.isNotBlack(1)测试结果：{}".format(s.isNotBlack(1)))
    except Exception as e:
        print("s.isNotBlack(1)测试结果：{}".format(str(e)))
    print("----------StrUtil测试------------\n")
    # 文件测试
    print("----------FileUtil测试------------")
    f = FileUtil()
    print('f.isExist("checkUtils1.py")测试结果：{}'.format(f.isExist("checkUtils1.py"))) 
    print('f.isExist("checkUtils.py")测试结果：{}'.format(f.isExist("checkUtils.py"))) 
    print("----------FileUtil测试------------\n")

def main():
    test()
    # input("按回车退出！")


if __name__ == '__main__':
    main()