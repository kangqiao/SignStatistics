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
