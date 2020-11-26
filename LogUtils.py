# -*- coding: utf-8 -*-

import logging.handlers
from ConfigUtils import Config
import os


class Log():
    '''日志'''
    __file_name = "log/log.txt"
    __cof_file_name = "config/config.txt"
    __logger = ""
    
    def __init__(self):
        self.__init_log()

    def __init__(self, conf_file_name):
        self.__cof_file_name = conf_file_name
        self.__init_log()

    def get_file_name(self):
        return self.__file_name

    def set_file_name(self, file_name):
        self.__file_name = file_name

    def __init_log(self):
        # 获取配置文件中的日志名
        c = Config()
        c.set_file_name(self.__cof_file_name)
        log_file_name = c.get_value("log_file_name")
        if not log_file_name:
            self.__file_name = log_file_name
        # 检查根目录创建
        file_dir = os.path.split(self.__file_name)[0]
        if not os.path.isdir(file_dir):
            os.makedirs(file_dir)
        # 清空日志文件
        if os.path.exists(self.__file_name):
            os.remove(self.__file_name)
        # 生成日志对象
        handler = logging.handlers.RotatingFileHandler(self.__file_name, maxBytes=1024 * 1024, backupCount=5, encoding='utf-8')  # 实例化handler
        fmt = '%(asctime)s - %(levelname)s - %(message)s'

        formatter = logging.Formatter(fmt)  # 实例化formatter
        handler.setFormatter(formatter)  # 为handler添加formatter

        self.__logger = logging.getLogger(self.__file_name[:self.__file_name.find(".")])  # 获取名为tst的logger
        self.__logger.addHandler(handler)  # 为logger添加handler
        self.__logger.setLevel(logging.DEBUG)

    def log_info(self, msg):
        self.__logger.info(msg)

    def log_error(self, msg):
        self.__logger.error(msg)

    def log_debug(self, msg):
        self.__logger.debug(msg)

    def log_warning(self, msg):
        self.__logger.warning(msg)


def main():
    log = Log("config/config.txt")
    print(log.get_file_name())
    log.log_info("异常！")


if __name__ == '__main__':
    main()

