import os
import re
import sys
from datetime import datetime

from openpyxl import Workbook

PRIMARY_SEPARATOR = r'[：|:]'
SECONDARY_SEPARATOR = r'[、|,|，|\s]'
REGEX_DATE = r'^(?P<year>\d{4})[-|年|\s]?(?P<month>\d{1,2})[-|月|\s]?(?P<day>\d{1,2})[-|日|\s]?[\s|（]*(?P<week>周[一|二|三|四|五|六|日])[）|\s]*$'

SIGN_TITLE = u'标题'
SIGN_DATE = u'奉粥日期'
SIGN_TEAM = u'团队'
SIGN_MANAGER = u'日负责人'
SIGN_SIGNER = u'签到'
SIGN_COOK_GRUEL = u'熬粥'
SIGN_COMPERE = u'前行'
SIGN_CUP_COUNT = u'杯数'
SIGN_PEOPLE_COUNT = u'人数'
SIGN_NEW_PEOPLE = u'新人数'
SIGN_PICTURE = u'摄影'
SIGN_DIARY = u'日志'
SIGN_PUBLICITY = u'文宣'
SIGN_FINISHED_MEET = u'结行'
SIGN_SUPPORTER = u'后勤'
SIGN_PROTECT_ENV = u'环保'
SIGN_SERVICE = u'奉粥'

STATISTICS_NAME = u'姓名'
STATISTICS_TOTAL_HOUR = u'总工时'

'''
熬粥2 日负责2 文宣2 后勤0.5 奉粥1 环保1
'''
Hour = {
    SIGN_COOK_GRUEL: 2,
    SIGN_MANAGER: 2,
    SIGN_PUBLICITY: 2,
    SIGN_SUPPORTER: 0.5,
    SIGN_SERVICE: 1,
    SIGN_PROTECT_ENV: 1
}


# 心栈每日签到表
class SignInfoDaily:

    def __init__(self):
        self.hourStatistics = {}

    def parseLine(self, lineData):
        if lineData and len(lineData) == 2:
            flag = lineData[0]
            info = lineData[1]
            if flag == SIGN_TITLE:
                self.title = info
            elif flag == SIGN_DATE:
                dateInfo = Utils.splitDateInfo(info)
                if dateInfo[0]:
                    self.date = dateInfo[0]
                if dateInfo[1]:
                    self.week = dateInfo[1]
            elif flag == SIGN_MANAGER:
                # 负责人记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_MANAGER)
                self.manager = info
            elif flag == SIGN_SIGNER:
                self.signer = info
            elif flag == SIGN_COOK_GRUEL:
                # 熬粥记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_COOK_GRUEL)
                self.cookGruel = info
            elif flag == SIGN_COMPERE:
                self.compere = info
            elif flag == SIGN_CUP_COUNT:
                self.cupCount = Utils.getFirstNumber(info)
            elif flag == SIGN_PEOPLE_COUNT:
                self.peopleCount = Utils.getFirstNumber(info)
            elif flag == SIGN_NEW_PEOPLE:
                self.newPeople = info
            elif flag == SIGN_PICTURE:
                self.picture = info
            elif flag == SIGN_DIARY:
                self.diary = info
            elif flag == SIGN_PUBLICITY:
                # 文宣记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_PUBLICITY)
                self.publicity = info
            elif flag == SIGN_FINISHED_MEET:
                self.finishedMeet = info
            elif flag == SIGN_SUPPORTER:
                # 后勤记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_SUPPORTER)
                self.supporter = info
            elif flag == SIGN_PROTECT_ENV:
                # 环保记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_PROTECT_ENV)
                self.protectEnv = info
            elif flag == SIGN_SERVICE:
                # 奉粥记录工时
                peopleList = Utils.splitPeople(info)
                self.recordHour(peopleList, SIGN_SERVICE)
                self.service = info

    def toSheetLineData(self):
        return [self.title, self.date, self.week, self.manager, self.signer, self.cookGruel, self.compere,
                self.cupCount, self.peopleCount, self.newPeople, self.picture, self.diary, self.publicity,
                self.finishedMeet, self.supporter, self.protectEnv, self.service]

    # 记录工时, 并保存每个人的工时到hourStatistics字典中<strName, Volunteer>{}
    def recordHour(self, peopleList, flag):
        for name in peopleList:
            if name not in self.hourStatistics:
                self.hourStatistics[name] = Volunteer(name, self.date, self.week)

            self.hourStatistics[name].setHourFlag(flag)


# 志愿者每日服务时常记录
class Volunteer:

    def __init__(self, name, date, week):
        self.name = name
        self.date = date
        self.week = week
        self.hourFlag = []

    def setHourFlag(self, flag):
        self.hourFlag.append(flag)

    def _desc(self, time, hour, content='奉粥'):
        return self.date + time + content + hour

    def genServiceRecord(self):
        totalHour = self.totalHour()
        if totalHour <= 0:
            return ""

        if totalHour >= 4:
            return self._desc("4:00~8:00", "4小时", "熬粥奉粥")
        elif totalHour >= 3.5:
            return self._desc("4:30~8:00", "3.5小时")
        elif totalHour >= 3:
            return self._desc("5:00~8:00", "3小时")
        elif totalHour >= 2.5:
            return self._desc("5:30~8:00", "2.5小时")
        elif totalHour >= 2:
            return self._desc("6:00~8:00", "2小时")
        elif totalHour >= 1.5:
            return self._desc("6:30~8:00", "1.5小时")
        elif totalHour >= 1:
            return self._desc("7:00~8:00", "1小时")
        elif totalHour >= 0.5:
            return self._desc("7:00~7:30", "0.5小时")
        else:
            return ""

    def totalHour(self):
        count = 0
        # 记录熬粥时常
        if SIGN_COOK_GRUEL in self.hourFlag:
            count += Hour[SIGN_COOK_GRUEL]

        # 同时间内优先顺序: 负责人 -> 文宣 -> 环保 -> 奉粥
        if SIGN_MANAGER in self.hourFlag:
            count += Hour[SIGN_MANAGER]
        elif SIGN_PUBLICITY in self.hourFlag:
            count += Hour[SIGN_PUBLICITY]
        elif SIGN_SUPPORTER in self.hourFlag and (SIGN_SERVICE in self.hourFlag or SIGN_PROTECT_ENV in self.hourFlag):
            count += Hour[SIGN_SUPPORTER] + Hour[SIGN_SERVICE]
        elif SIGN_PROTECT_ENV in self.hourFlag:
            count += Hour[SIGN_PROTECT_ENV]
        elif SIGN_SERVICE in self.hourFlag:
            count += Hour[SIGN_SERVICE]

        return count


class SignInfoSheet:
    SHEET_SIGN_DAILY_TITLE = [SIGN_TITLE, SIGN_DATE, SIGN_TEAM, SIGN_MANAGER, SIGN_SIGNER, SIGN_COOK_GRUEL,
                              SIGN_COMPERE,
                              SIGN_CUP_COUNT, SIGN_PEOPLE_COUNT, SIGN_NEW_PEOPLE, SIGN_PICTURE, SIGN_DIARY,
                              SIGN_PUBLICITY,
                              SIGN_FINISHED_MEET, SIGN_SUPPORTER, SIGN_PROTECT_ENV, SIGN_SERVICE]

    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb.create_sheet(title=u'签到表', index=0)
        self.sheet.append(SignInfoSheet.SHEET_SIGN_DAILY_TITLE)
        # <日期, 行数据> 记录每天的签到信息到一行中.
        self.signInfoDaily = {}

    def appendDailyInfo(self, daily):
        self.signInfoDaily[daily.date] = daily.toSheetLineData()

    def writeSheet(self):
        if self.signInfoDaily:
            _signInfoDaily = sorted(self.signInfoDaily.items(), key=lambda d: Utils.parseDate(d[0]), reverse=True)
            for date, line in _signInfoDaily:
                self.sheet.append(line)


class StatisticsSheet:
    SHEET_DEFAULT_TITLE = [STATISTICS_NAME, STATISTICS_TOTAL_HOUR]

    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb.create_sheet(title=u'工时统计表', index=1)
        # 初始化两列数据[ 姓名, 总工时 ]
        self.sheet.append(StatisticsSheet.SHEET_DEFAULT_TITLE)
        # <名字, 总工时> 的统计信息字典
        self.nameTotalHour = {}
        # <名字, <日期, 工时> 的统计信息字典
        self.nameDateHour = {}
        # <日期, 图表的列标示> 的统计信息
        self.dateList = []

    def appendDailyStatistics(self, daily):
        if daily.date not in self.dateList:
            self.dateList.append(daily.date)

        self.dateList.sort()

        for name in daily.hourStatistics:
            if name not in self.nameTotalHour:
                self.nameTotalHour[name] = 0

            if name not in self.nameDateHour:
                # 为某人在Sheet表中建立<日期, 工时>的映射字典
                self.nameDateHour[name] = {}

            dailyHour = daily.hourStatistics[name].totalHour()
            # 保存某人当天的工时
            self.nameDateHour[name][daily.date] = dailyHour
            # 记录某人的所有工时
            self.nameTotalHour[name] += dailyHour

    # 最后写入Excel表中
    def writeSheet(self):
        if not self.nameDateHour : return

        # 按人员总工时排序
        _nameTotalHour = sorted(self.nameTotalHour.items(), key=lambda d: int(d[1]), reverse=True)

        _dateList = sorted(self.dateList, key=lambda x: Utils.parseDate(x), reverse=True)
        for col, date in enumerate(_dateList, len(StatisticsSheet.SHEET_DEFAULT_TITLE) + 1):
            cell = self.sheet.cell(row=1, column=col)
            cell.value = date

        row = 1
        for name, totalHour in _nameTotalHour:
            row += 1

            cell = self.sheet.cell(row=row, column=1)
            cell.value = name
            cell = self.sheet.cell(row=row, column=2)
            cell.value = totalHour

            for column, date in enumerate(_dateList, len(StatisticsSheet.SHEET_DEFAULT_TITLE) + 1):
                if date in self.nameDateHour[name]:
                    cell = self.sheet.cell(row=row, column=column)
                    cell.value = self.nameDateHour[name][date]


class VolunteerBJSheet:
    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb.create_sheet(title=u'志愿北京', index=2)
        self.sheet.append(['姓名', '志愿北京记录'])
        # <名字, 数据记录>
        self.volunteerBJRecord = {}
        # <名字, 总工时> 的统计信息字典
        self.nameTotalHour = {}

    def appendVolunteerBJRecord(self, daily):
        for name in daily.hourStatistics:
            if name not in self.volunteerBJRecord:
                self.volunteerBJRecord[name] = ""

            volunteer = daily.hourStatistics[name]
            self.volunteerBJRecord[name] += volunteer.genServiceRecord() + "\n"

    def syncNameTotalHourData(self, nameTotalHour):
        self.nameTotalHour = nameTotalHour

    def writeSheet(self):
        if not self.volunteerBJRecord : return

        # 按人员总工时排序
        _nameTotalHour = sorted(self.nameTotalHour.items(), key=lambda d: int(d[1]), reverse=True)

        row = 1
        for name, totalHour in _nameTotalHour:
            row += 1
            cell = self.sheet.cell(row=row, column=1)
            cell.value = name
            recordContent = name + " - 共" + str(self.nameTotalHour[name]) + "小时\n" + self.volunteerBJRecord[name]
            cell = self.sheet.cell(row=row, column=2)
            cell.value = recordContent


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
    def splitPeople(str):
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


wb = Workbook()
signInfoSheet = SignInfoSheet(wb)
statisticsSheet = StatisticsSheet(wb)
volunteerBJSheet = VolunteerBJSheet(wb)


def parseServiceHourByFile(filePath):
    with open(filePath) as file:
        daily = SignInfoDaily()
        for line in file:

            # 判断是否空行
            line = line.strip()
            if not line:
                continue

            # 分隔主要数据
            primary_data = Utils.splitLine(line)

            # 检查数据
            if not primary_data or len(primary_data) != 2:
                Utils.log("解析行数据失败>>>" + line)
                continue

            if primary_data[0] == SIGN_TITLE:
                # 遇到 "标题" 初始化操作
                daily = SignInfoDaily()

            daily.parseLine(primary_data)

            if primary_data[0] == SIGN_SERVICE:
                # 遇到 "奉粥" 当天日志整理完毕, 写入对应Excel表格中.
                signInfoSheet.appendDailyInfo(daily)
                statisticsSheet.appendDailyStatistics(daily)
                volunteerBJSheet.appendVolunteerBJRecord(daily)

        signInfoSheet.writeSheet()
        statisticsSheet.writeSheet()
        volunteerBJSheet.syncNameTotalHourData(statisticsSheet.nameTotalHour)
        volunteerBJSheet.writeSheet()


def parseServiceHourByPath(dirPath):
    files = os.listdir(dirPath)
    for file in files:
        if not os.path.isdir(file) and os.path.splitext(file)[-1] == '.txt':
            parseServiceHourByFile(file)

    wb.save("签到统计表.xlsx")


if __name__ == '__main__':
    path = sys.argv[1:]
    if path[0]:
        print(path[0])
        parseServiceHourByPath(path[0])
    else:
        print("请指定签到信息文件或目录")

    #parseServiceHourByPath(".")
