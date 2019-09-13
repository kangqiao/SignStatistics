# -*- coding: UTF-8 -*-
import os
import re

from datetime import datetime

PRIMARY_SEPARATOR = r'[：|:]'
MAPPING_SEPARATOR = r'[=]'
SECONDARY_SEPARATOR = r'[、|,|，|\s]'
REGEX_DATE = r'^(?P<year>\d{4})[-|年|\s]?(?P<month>\d{1,2})[-|月|\s]?(?P<day>\d{1,2})[-|日|\s]?[\s|（]*(?P<week>周[一|二|三|四|五|六|日])[）|\s]*$'

class Utils:

    @staticmethod
    def log(str):
        print(str)

    @staticmethod
    def getFirstNumber(str):
        if str:
            num_arr = re.findall(r'\d+', str)
            if num_arr:
                return num_arr[0]
        return 0

    @staticmethod
    def splitLine(str):
        if str:
            return re.split(PRIMARY_SEPARATOR, str)
        return []

    @staticmethod
    def split(str, separator=MAPPING_SEPARATOR):
        if str:
            return re.split(separator, str)
        return []

    @staticmethod
    def splitName(str):
        if str:
            return re.split(SECONDARY_SEPARATOR, str)
        return []

    @staticmethod
    def splitDateInfo(str):
        mat = re.search(REGEX_DATE, str.strip())
        if mat:
            # return (datetime(int(mat.group('year')), int(mat.group('month')), int(mat.group('day'))), mat.group("week"))
            return (mat.group('year') + '年' + mat.group('month') + '月' + mat.group('day') + '日', mat.group("week"))

    @staticmethod
    def parseDate(date):
        if date:
            return datetime.strptime(date, "%Y年%m月%d日")

    @staticmethod
    def formatDate(date, week):
        if date:
            return date.strftime("%Y年%m月%d日") + ' ' + week

    @staticmethod
    def isTxtFile(path):
        return os.path.isfile(path) and os.path.splitext(path)[-1] == '.txt'
