# -*- coding: UTF-8 -*-
from utils import *

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
STATISTICS_TOTAL_DAYS = u'总天数'
STATISTICS_AVERAGE_DAYS = u'日均'

'''
熬粥3 日负责2 文宣3 后勤0.5 奉粥1 环保1 结行0.5
'''
Hour = {
    SIGN_COOK_GRUEL: 3,
    SIGN_MANAGER: 2,
    SIGN_PUBLICITY: 3,
    SIGN_SUPPORTER: 0.5,
    SIGN_SERVICE: 1,
    SIGN_PROTECT_ENV: 1,
    SIGN_FINISHED_MEET: 0.5
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
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_MANAGER)
                self.manager = info
            elif flag == SIGN_SIGNER:
                self.signer = info
            elif flag == SIGN_COOK_GRUEL:
                # 熬粥记录工时
                peopleList = Utils.splitName(info)
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
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_PUBLICITY)
                self.publicity = info
            elif flag == SIGN_FINISHED_MEET:
                # 结行记录工时
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_FINISHED_MEET)
                self.finishedMeet = info
            elif flag == SIGN_SUPPORTER:
                # 后勤记录工时
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_SUPPORTER)
                self.supporter = info
            elif flag == SIGN_PROTECT_ENV:
                # 环保记录工时
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_PROTECT_ENV)
                self.protectEnv = info
            elif flag == SIGN_SERVICE:
                # 奉粥记录工时
                peopleList = Utils.splitName(info)
                self.recordHour(peopleList, SIGN_SERVICE)
                self.service = info

    def toSheetLineData(self):
        return [self.title, self.date, self.week, self.manager, self.signer, self.cookGruel, self.compere,
                self.cupCount, self.peopleCount, self.newPeople, self.picture, self.diary, self.publicity,
                self.finishedMeet, self.supporter, self.protectEnv, self.service]

    # 记录工时, 并保存每个人的工时到hourStatistics字典中<strName, Volunteer>{}
    def recordHour(self, peopleList, flag):
        for name in peopleList:
            user = UserManager.findUser(name)
            if user not in self.hourStatistics:
                self.hourStatistics[user] = Volunteer(user, self.date, self.week)

            self.hourStatistics[user].setHourFlag(flag)


# 志愿者每日服务时常记录
class Volunteer:

    def __init__(self, user, date, week):
        self.user = user
        self.date = date
        self.week = week
        self.hourFlag = []

    @property
    def name(self):
        return self.user.name

    def setHourFlag(self, flag):
        self.hourFlag.append(flag)

    def _desc(self, time, hour, content='奉粥'):
        return self.date + time + content + hour

    def genServiceRecord(self):
        totalHour = self.totalHour()
        if totalHour <= 0:
            return ""

        if totalHour >= 4:
            return self._desc(" ", str(totalHour)+"小时", "熬粥奉粥")
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
        elif SIGN_PUBLICITY in self.hourFlag:
            count += Hour[SIGN_PUBLICITY]

        # 同时间内优先顺序: 负责人 -> 文宣 -> 环保 -> 奉粥
        if SIGN_MANAGER in self.hourFlag:
            count += Hour[SIGN_MANAGER]
        # elif SIGN_PUBLICITY in self.hourFlag:
        #     count += Hour[SIGN_PUBLICITY]
        elif SIGN_SUPPORTER in self.hourFlag and (SIGN_SERVICE in self.hourFlag or SIGN_PROTECT_ENV in self.hourFlag):
            count += Hour[SIGN_SUPPORTER] + Hour[SIGN_SERVICE]
        elif SIGN_PROTECT_ENV in self.hourFlag:
            count += Hour[SIGN_PROTECT_ENV]
        elif SIGN_SERVICE in self.hourFlag:
            count += Hour[SIGN_SERVICE]

        # 结行单独增加0.5工时
        if SIGN_FINISHED_MEET in self.hourFlag:
            count += Hour[SIGN_FINISHED_MEET]
        return count

class User:

    def __init__(self, id, real, other=''):
        self.userId = id
        self.realName = real
        self.otherName = other

    @property
    def name(self):
        if self.realName:
            return self.realName
        else:
            return self.otherName

    @property
    def id(self):
        if self.userId:
            return self.userId
        else:
            return "未注册"

    def __repr__(self):
        if self.userId:
            return self.userId
        if self.realName:
            return self.realName
        else:
            return self.otherName

    def __str__(self):
        return self.__repr__()


class UserManager:
    # 志愿北京信息<ID, (真实姓名, 别名1, 别名2...)>
    nameMapping = {}
    nameUserMap = {}

    @staticmethod
    def setNameMapping(mapping):
        UserManager.nameMapping = mapping

    @staticmethod
    def findNameMapping(name):
        # 返回(ID, 真实姓名)
        if not name:
            return (None, None)

        for id, list in UserManager.nameMapping.items():
            for val in list:
                if name == val:
                    return (id, list[0])

        return (None, None)

    @staticmethod
    def findUser(name):
        # 根据name去志愿北京中查找,
        id, realName = UserManager.findNameMapping(name)
        if id:
            if id in UserManager.nameUserMap:
                return UserManager.nameUserMap[id]
            else:
                user = User(id, realName, other=name)
                UserManager.nameUserMap[id] = user
                return user
        else:
            if name in UserManager.nameUserMap:
                return UserManager.nameUserMap[name]
            else:
                user = User(id, realName, other=name)
                UserManager.nameUserMap[user.name] = user
                return user
