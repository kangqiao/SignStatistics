# -*- coding: UTF-8 -*-
import os
import sys

from openpyxl import Workbook

from sheet import *
from utils import *

wb = Workbook()
signInfoSheet = SignInfoSheet(wb)
statisticsSheet = StatisticsSheet(wb)
volunteerBJSheet = VolunteerBJSheet(wb)
topSheet = TopSheet(wb)

def loadMappingConifg(filePath):
    with open(filePath, 'r', encoding='utf-8') as file:
        mapping_data = {}
        for line in file:

            line = line.encode('utf-8').decode('utf-8-sig').strip()

            line = line.strip()
            if not line:
                continue

            if line.startswith('#'):
                continue

            data = Utils.split(line)

            if data and len(data) == 2:
                mapping_data[data[0]] = tuple(Utils.splitName(data[1]))

        # 初始化用户信息列表映射
        UserManager.setNameMapping(mapping_data)


def parseServiceHourByFile(filePath):
    with open(filePath, 'r', encoding='utf-8') as file:
        daily = SignInfoDaily()
        for line in file:

            # python读写文件出现\ufeff
            # question: https://blog.csdn.net/boystray/article/details/80714347
            # solution: https://blog.csdn.net/xiazhipeng1000/article/details/79720391
            line = line.encode('utf-8').decode('utf-8-sig').strip()

            # 判断是否空行
            line = line.strip()
            if not line or line == '':
                continue

            # 分隔主要数据
            primary_data = Utils.splitLine(line)

            # 检查数据
            if not primary_data or len(primary_data) != 2:
                Utils.log("解析行数据失败>>> filePath" + filePath + ', ' + primary_data)
                continue

            if primary_data[0] == SIGN_TITLE:
                # 遇到 "标题" 初始化操作
                daily = SignInfoDaily()

            daily.parseLine(primary_data)

            if primary_data[0] == SIGN_SERVICE:
                # 遇到 "奉粥" 当天日志整理完毕, 追加数据到对应Excel表缓存中.
                signInfoSheet.appendDailyInfo(daily)
                statisticsSheet.appendDailyStatistics(daily)
                volunteerBJSheet.appendVolunteerBJRecord(daily)
                topSheet.appendTopInfo(daily)


def writeSheet():
    signInfoSheet.writeSheet()
    statisticsSheet.writeSheet()
    volunteerBJSheet.syncNameTotalHourData(statisticsSheet.nameTotalHour)
    volunteerBJSheet.writeSheet()
    topSheet.writeSheet()


def parseServiceHourByPath(path, outputFile):
    fileList = []
    if os.path.isdir(path):
        # 过滤txt文件
        fileList = [os.path.join(path, p) for p in os.listdir(path) if Utils.isTxtFile(os.path.join(path, p))]
    elif Utils.isTxtFile(path):
        fileList.append(path)

    if os.path.isfile(MAPPING_CONFIG):
        loadMappingConifg(MAPPING_CONFIG)

    for file in fileList:
        parseServiceHourByFile(file)

    if fileList:
        writeSheet()
        wb.save(outputFile+".xlsx")

MAPPING_CONFIG = './mapping.ini'
CMD_PARAM_PATH = 'path'
CMD_PARAM_OUTPUT = "output"

if __name__ == '__main__':
    paramList = sys.argv[1:]
    outputFile = "签到统计表"
    path = ""
    for param in paramList:
        data = Utils.split(param)
        if data and data[0] == CMD_PARAM_PATH:
            path = data[1]
        elif data and data[0] == CMD_PARAM_OUTPUT:
            outputFile = data[1]

    print("path:" + path)
    print("outputFile:" + outputFile)

    # path = "/Users/zhaopan/workspace/python/service_hour/2019/5月签到表.txt"
    if path:
        parseServiceHourByPath(path, outputFile)
    else:
        print("请指定签到信息文件或目录")