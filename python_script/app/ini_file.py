# -*- coding: utf-8 -*-
import configparser
import os
from logging_err import *


class open_ini_file():
    __instance = None

    @staticmethod
    def inst():
        if open_ini_file.__instance == None:
            open_ini_file.__instance = open_ini_file()
        return open_ini_file.__instance

    def __init__(self):
        import sys, os
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the pyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            self.BASE_DIR = os.getcwd()
        else:
            self.BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.ini_parser = configparser.ConfigParser()
        self.path_ini = os.path.join(self.BASE_DIR, 'con_config.ini')
        self.ini_parser.read(self.path_ini)
        # try:
        #     # self.ini_parser = configparser.ConfigParser()
        #     # self.BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        #     #
        #     # self.path_ini = os.path.join(self.BASE_DIR, 'con_config.ini')
        #
        #     # self.__class__.ini_parser.read(self.path_ini)
        # except Exception as e:
        #     exeption_print(e)

    def __add(self, path, station):
        print(path)
        self.ini_parser.set('FILE_PROJECT', 'path', os.path.join(self.BASE_DIR, os.path.join('project', path)))
        self.ini_parser.set('STATION', '0', station)
        with open(self.path_ini, 'wt') as con_config:
            self.ini_parser.write(con_config)

    def __open_config_file(self, section):
        for get_option in self.ini_parser.options(section):
            out = self.ini_parser.get(section, get_option)
            yield out

    def get_path(self):
        path = self.ini_parser.get('FILE_PROJECT', 'path')
        return path

    def get_STATION(self):
        STATION = []
        for conv_file in self.__open_config_file('STATION'):
            STATION.append(conv_file)
        return STATION

    def get_IN_PARAM(self):
        IN_PARAM = []
        for in_p in self.__open_config_file('IN_PARAM'):
            IN_PARAM.append(in_p)
        return IN_PARAM

    def set_path_and_station(self,path,station):
        self.__add(path,station)