# coding:utf-8
# @Time     : 2023/7/4 18:13
# @Author   : vicky 
# @Email    : 530622716@qq.com
# @File     ：DeviceInfo.py


import os
import time

class Projectorinfo:
  #  def __init__(self):

#获取设备的基础信息,如厂家，型号，系统版本
    def devicesinfo(self):
        deviceName=os.popen("adb shell getprop ro.product.model").read()
        platformVersion=os.popen("adb shell getprop ro.build.version.release").read()
        producer=os.popen("adb shell getprop ro.product.brand").read()
        return "产品型号：%s %s,系统版本：Android %s " % (
        producer.replace("\n"," "),deviceName.replace("\n"," "),platformVersion.replace("\n"," ")
        )

#检测设备是否连接成功，连接成功返回成功，连接失败返回失败
    def check_devices(self):
        try:
            devices=os.popen("adb devices").read()
            if "device" in devices.split("\n")[1]:
                print("V"*20,"连接成功","K"*20)
                print(self.devicesinfo())
                return True
            else:
                print("K"*20,"未连接上设备，请检测设备连接","V"*20)
        except Exception as e:
            print("设备连接不成功，报错：",e)

Mydeveice=Projectorinfo()
Mydeveice.check_devices()
