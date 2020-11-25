# -*- coding:utf-8 -*-
##
#
#   配置文件读取工具
#
##
import io
import sys
import os
import time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class Config:
    """配置文件对象"""
    __file_name = ''  # 配置文件名

    def __init__(self):
        self.__file_name = "config.txt"

    def get_file_name(self):
        return self.__file_name

    def set_file_name(self, file_name):
        self.__file_name = file_name

    # 获取所有key
    def get_keys(self):
        contents = []
        file_contents = self.get_file_contents()
        if file_contents:
            for content in file_contents:
                if content.find("=") is not -1:
                    contents.append(content[:content.find(":")].strip())
        return contents

    # 获取所有配置信息
    def get_file_contents(self):
        contents = []
        if self.__file_is_exist():
            with open(self.__file_name, "r+", encoding="utf-8") as configFile:
                content = configFile.read()
                contents = content.split("\n")
        return contents

    # 获取所有值
    def get_value(self, key):
        # 获取配置参数
        file_list = self.get_file_contents()
        for x in file_list:
            x = x.strip()
            if x.find(key) is not -1:
                value = x[x.find("=")+1:]
                return value.strip()

    # 添加新属性，有则覆盖
    def add_value(self, key, value):
        exist_flag = False
        # 获取配置参数
        file_list = self.get_file_contents()
        for i, val in enumerate(file_list):
            x = val.strip()
            if x.find(key) is not -1:
                file_list[i] = key + "=" + value
                exist_flag = True
                break
        if exist_flag is not True:
            file_list.append(key + "=" + value)
        contents = "\n".join(file_list)
        with open(self.__file_name, "w", encoding="utf-8") as f:
            f.write(contents)

    # 判断配置文件是否存在
    def __file_is_exist(self):
        if os.path.exists(self.__file_name):
            return True
        return False


def main():
    c = Config()
    print(c.get_value("name"))
    time.sleep(2)


if __name__ == '__main__':
    main()
