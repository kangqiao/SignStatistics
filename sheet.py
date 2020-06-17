# -*- coding: UTF-8 -*-
from model import *
from utils import *


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
    SHEET_DEFAULT_TITLE = [STATISTICS_NAME, STATISTICS_TOTAL_HOUR, STATISTICS_TOTAL_DAYS, STATISTICS_AVERAGE_DAYS]

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

        for user in daily.hourStatistics:
            if user not in self.nameTotalHour:
                self.nameTotalHour[user] = 0

            if user not in self.nameDateHour:
                # 为某人在Sheet表中建立<日期, 工时>的映射字典
                self.nameDateHour[user] = {}

            dailyHour = daily.hourStatistics[user].totalHour()
            # 保存某人当天的工时
            self.nameDateHour[user][daily.date] = dailyHour
            # 记录某人的所有工时
            self.nameTotalHour[user] += dailyHour

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
        for user, totalHour in _nameTotalHour:
            row += 1

            cell = self.sheet.cell(row=row, column=1)
            cell.value = user.name
            cell = self.sheet.cell(row=row, column=2)
            cell.value = totalHour
            totalDays = len(self.nameDateHour[user])
            cell = self.sheet.cell(row=row, column=3)
            cell.value = totalDays
            cell = self.sheet.cell(row=row, column=4)
            cell.value = totalHour / totalDays

            for column, date in enumerate(_dateList, len(StatisticsSheet.SHEET_DEFAULT_TITLE) + 1):
                if date in self.nameDateHour[user]:
                    cell = self.sheet.cell(row=row, column=column)
                    cell.value = self.nameDateHour[user][date]


class VolunteerBJSheet:
    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb.create_sheet(title=u'志愿北京', index=2)
        self.sheet.append(['ID', '姓名', '志愿北京记录'])
        # <名字, 数据记录>
        self.volunteerBJRecord = {}
        # <名字, 总工时> 的统计信息字典
        self.nameTotalHour = {}

    def appendVolunteerBJRecord(self, daily):
        for user in daily.hourStatistics:
            if user not in self.volunteerBJRecord:
                self.volunteerBJRecord[user] = ""

            volunteer = daily.hourStatistics[user]
            self.volunteerBJRecord[user] += volunteer.genServiceRecord() + "\n"

    def syncNameTotalHourData(self, nameTotalHour):
        self.nameTotalHour = nameTotalHour

    def writeSheet(self):
        if not self.volunteerBJRecord : return

        # 按人员总工时排序
        _nameTotalHour = sorted(self.nameTotalHour.items(), key=lambda d: int(d[1]), reverse=True)

        row = 1
        for user, totalHour in _nameTotalHour:
            row += 1
            cell = self.sheet.cell(row=row, column=1)
            cell.value = user.id
            cell = self.sheet.cell(row=row, column=2)
            cell.value = user.name
            recordContent = user.name + " - 共" + str(self.nameTotalHour[user]) + "小时\n" + self.volunteerBJRecord[user]
            cell = self.sheet.cell(row=row, column=3)
            cell.value = recordContent


class TopSheet:
    SHEET_DEFAULT_TITLE = [STATISTICS_NAME, SIGN_SERVICE, SIGN_MANAGER, SIGN_SIGNER, SIGN_COOK_GRUEL, SIGN_COMPERE, SIGN_PICTURE, SIGN_DIARY, SIGN_PUBLICITY, SIGN_SUPPORTER, SIGN_PROTECT_ENV]

    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb.create_sheet(title=u'Top排名', index=3)
        self.sheet.append(TopSheet.SHEET_DEFAULT_TITLE)
        # <name, {总天数: 1, 签名: 4, 熬粥: 5, 摄影: 5, 日志: 5, ....}>
        self.topInfo = {}

    def appendTopInfo(self, daily):
        if daily.service:
            self.statisticsTop(Utils.splitName(daily.service), SIGN_SERVICE)
            # 由于环保不会加到奉粥中, 所以此处统计奉粥时将环保也加入其中
            self.statisticsTop(Utils.splitName(daily.protectEnv), SIGN_SERVICE)
        if daily.manager:
            self.statisticsTop(Utils.splitName(daily.manager), SIGN_MANAGER)
        if daily.signer:
            self.statisticsTop(Utils.splitName(daily.signer), SIGN_SIGNER)
        if daily.cookGruel:
            self.statisticsTop(Utils.splitName(daily.cookGruel), SIGN_COOK_GRUEL)
        if daily.compere:
            self.statisticsTop(Utils.splitName(daily.compere), SIGN_COMPERE)
        if daily.picture:
            self.statisticsTop(Utils.splitName(daily.picture), SIGN_PICTURE)
        if daily.diary:
            self.statisticsTop(Utils.splitName(daily.diary), SIGN_DIARY)
        if daily.publicity:
            self.statisticsTop(Utils.splitName(daily.publicity), SIGN_PUBLICITY)
        if daily.supporter:
            self.statisticsTop(Utils.splitName(daily.supporter), SIGN_SUPPORTER)
        if daily.protectEnv:
            self.statisticsTop(Utils.splitName(daily.protectEnv), SIGN_PROTECT_ENV)


    def statisticsTop(self, peopleList, type):
        if peopleList and len(peopleList) > 0:
            for people in peopleList:
                id, name = UserManager.findNameMapping(people)
                if not name:
                    name = people

                if name not in self.topInfo:
                    self.topInfo[name] = {}
                    # 依据的排序key, 所以默认赋值为0
                    self.topInfo[name][SIGN_SERVICE] = 0

                if type not in self.topInfo[name]:
                    self.topInfo[name][type] = 1
                else:
                    self.topInfo[name][type] += 1

    def writeSheet(self):
        if not self.topInfo:
            return

        # 按人员总工时排序
        _sortedTopInfo = sorted(self.topInfo.items(), key=lambda item: int(item[1][SIGN_SERVICE]), reverse=True)

        row = 1
        for name, info in _sortedTopInfo:
            if name == u'无':
                continue
            row += 1
            for i, val in enumerate(TopSheet.SHEET_DEFAULT_TITLE):
                if val == STATISTICS_NAME:
                    cell = self.sheet.cell(row=row, column=i+1)
                    cell.value = name
                    continue

                if val not in info:
                    info[val] = 0
                cell = self.sheet.cell(row=row, column=i+1)
                cell.value = info[val]