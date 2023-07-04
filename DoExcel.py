# coding:utf-8
# @Time     : 2023/7/4 18:36
# @Author   : vicky 
# @Email    : 530622716@qq.com
# @File     ：DoExcel.py
#脚本对Excel进行读写操作，可以读取excel中的表格数据，数据写入到excel
import xlwt
import xlrd
from openpyxl import load_workbook#支持对excel读写，支持.xlsx后缀，不支持.xls
import os
class DoExcel:

    def get_path(self):
        file_path=os.path.join(os.path.dirname(os.path.abspath(__file__)),"test.xlsx")#获取绝对路径
        return file_path

    def read_excel(self):
        wb=load_workbook(self.get_path()) #打开excel


    read_excel("E:\\example.xlsx")


    def write_excel(self):
        xxx
        return

