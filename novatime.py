# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL
# pylint: disable=logging-fstring-interpolation

"""Wrapper for the "REST API" for NOVATime"""

import configparser
import json
import logging
import os
from dataclasses import dataclass
from datetime import timedelta
from os.path import isdir, isfile, join
from typing import Any, Dict, List

import arrow
import requests
from bs4 import BeautifulSoup
from rich.console import Console, ConsoleOptions, RenderResult
from rich.logging import RichHandler
from rich.progress import Progress
from rich.table import Table

PARAM_DATE_FORMAT = 'ddd MMM DD YYYY'
PUNCH_FORMAT = 'MM/DD/YYYY HH:mm:ss'
PUNCH_TIME_FORMAT = 'HH:mmA'
NOVA_DATE_FORMAT = 'M/D/YYYY'
PUNCH_DATEKEY_FORMAT = 'MM/DD'


FORMAT = "%(message)s"
logging.basicConfig(
    level="DEBUG",
    format=FORMAT,
    datefmt="[%X]",
    handlers=[RichHandler(rich_tracebacks=True,
                          tracebacks_suppress=[requests])]
)
logging.getLogger("requests").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)


def parse_punch(punch_str):
    """Parse a punch entry into an `arrow`"""
    return arrow.get(punch_str, PUNCH_FORMAT).replace(tzinfo='America/Detroit')


def parse_date(date_str):
    """Parse a date entry into an `arrow`"""
    return arrow.get(date_str, NOVA_DATE_FORMAT).replace(tzinfo='America/Detroit')


@dataclass
class DatePeriod:
    """Pay period with start and end dates."""
    start: arrow.arrow.Arrow
    end: arrow.arrow.Arrow

    def __str__(self) -> str:
        return f'{self.start.date()}--{self.end.date()}'

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]PayPeriod:[/b] {self.start} to {self.end}"


@dataclass
class GPS:
    """GPS coordinates."""
    latitude: float
    longitude: float

    def __str__(self) -> str:
        return f'{self.latitude}, {self.longitude}'

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]GPS:[/b] {self.latitude}, {self.longitude}"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "nGPSLatitude": self.latitude,
            "nGPSLongitude": self.longitude
        }


@dataclass
class User:
    username: str
    password: str
    user_seq: str
    employee_seq: str
    access_seq: str
    first_name: str
    last_name: str
    full_name: str

    def __init__(self, username: str,
                 password: str,
                 user_seq: str = None,
                 employee_seq: str = None,
                 access_seq: str = None,
                 first_name: str = None,
                 last_name: str = None,
                 full_name: str = None) -> None:
        self.username = username
        self.password = password
        self.user_seq = user_seq
        self.employee_seq = employee_seq
        self.access_seq = access_seq
        self.first_name = first_name
        self.last_name = last_name
        self.full_name = full_name

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]User:[/b] #{self.full_name} ({self.username})"
        my_table = Table("Attribute", "Value")
        my_table.add_row("password", self.password)
        my_table.add_row("user_seq", self.user_seq)
        my_table.add_row("employee_seq", self.employee_seq)
        my_table.add_row("access_seq", self.access_seq)
        my_table.add_row("first_name", self.first_name)
        my_table.add_row("last_name", self.last_name)
        yield my_table


@dataclass
class HourTotals:
    last_week: timedelta
    this_week: timedelta
    total: timedelta

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]HourTotals:[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("last_week", str(self.last_week))
        my_table.add_row("this_week", str(self.this_week))
        my_table.add_row("total", str(self.total))
        yield my_table


@dataclass
class EntryCategory:
    """Representation of an entry category in NOVATime."""
    # pylint: disable=too-many-instance-attributes
    # There is no good way to package these into fewer attributes, so pylint can blow it
    value: int
    description: str
    group_number: int
    group_seq: int
    group_code: str
    group_user_type: int
    gps: GPS
    group_color: str
    group_value_description: str
    closed: bool

    def __init__(self, group: Dict[str, Any]) -> None:
        self.description = group['cDescription']
        self.group_code = group['cGroupCode']
        self.group_color = group['cGroupColor']
        self.value = group['cGroupValue']
        self.group_value_description = group['cGroupValueDescription']
        self.group_number = group['iGroupNumber']
        self.group_user_type = group['iGroupUserType']
        self.group_seq = group['iGroupValueSeq']
        self.closed = group['lClosed']
        self.gps = GPS(latitude=group['nGPSLatitude'],
                       longitude=group['nGPSLongitude'])

    def __str__(self) -> str:
        return f'{self.value} [{self.description or self.group_value_description}]'

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "iGroupNumber": self.group_number,
            "iGroupValueSeq": self.group_seq,
            "cGroupValue": str(self.value),
            "cGroupValueDescription": self.group_value_description,
            "cGroupCode": self.group_code,
            "iGroupUserType": self.group_user_type,
            "cGroupColor": self.group_color,
            "cDescription": self.description,
            "lClosed": self.closed
        }
        d.update(self.gps.to_dict())
        return d

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]EntryCategory:[/b] #{self.group_value_description}"
        my_table = Table("Attribute", "Value")
        my_table.add_row("value", str(self.value))
        my_table.add_row("description", self.description)
        my_table.add_row("group_number", str(self.group_number))
        my_table.add_row("group_seq", str(self.group_seq))
        my_table.add_row("group_code", self.group_code)
        my_table.add_row("group_user_type", str(self.group_user_type))
        my_table.add_row("gps", self.gps)
        my_table.add_row("group_color", self.group_color)
        my_table.add_row("closed", str(self.closed))
        yield my_table


@dataclass
class EntryCategoryGroup:
    """Representation of a group of entry categories in NOVATime."""
    value: int
    name: str
    options: Dict[str, EntryCategory]

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]EntryCategoryGroup:[/b] {self.value} [{self.name}]"
        my_table = Table('value',
                         'group_value_description',
                         'description',
                         'group_number',
                         'group_seq',
                         'group_code',
                         'group_user_type',
                         'gps',
                         'group_color',
                         'closed')
        for _, option in self.options.items():
            my_table.add_row(str(option.value),
                             option.group_value_description,
                             option.description,
                             str(option.group_number),
                             str(option.group_seq),
                             option.group_code,
                             str(option.group_user_type),
                             str(option.gps),
                             option.group_color,
                             str(option.closed))
        yield my_table


@dataclass
class PayCode:
    """Representation of the pay code fields in a NOVATime timesheet."""
    code: int           # nPayCode
    description: str    # cExpCode
    code_type: int      # nCodeType
    read_only: bool     # lPayCodeReadOnly
    pay_type: str       # cPayType
    policy: int         # nPayPolicy
    policy_description: str  # cPayPolicyDescription
    rate: float         # nPayRate

    def __str__(self) -> str:
        return f'{self.code}[{self.description}]'

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]PayCode:[/b] #{self.code}"
        my_table = Table("Attribute", "Value")
        my_table.add_row("description", self.description)
        my_table.add_row("code_type", str(self.code_type))
        my_table.add_row("read_only", str(self.read_only))
        my_table.add_row("pay_type", self.pay_type)
        my_table.add_row("policy", str(self.policy))
        my_table.add_row("policy_description", self.policy_description)
        my_table.add_row("rate", str(self.rate))
        yield my_table


@dataclass
class TimesheetEntryException:
    """Representation of the exception fields in a NOVATime timesheet."""
    # pylint: disable=too-many-instance-attributes
    # There is no good way to package these into fewer attributes, so pylint can blow it
    overtime: bool
    tardy: bool
    early_out: bool
    early_in: bool
    late_out: bool
    missing_punch: bool
    meal_break_premium: bool
    meal_desc: str
    under_pay: bool
    over_pay: bool
    tardy_grace: bool
    early_out_grace: bool
    unpaid_break: bool
    break_desc: str
    auto_deduct_meal: bool
    auto_meal_waived: bool
    unconfirmed_punch: bool
    unconfirmed_in_punch: bool
    unconfirmed_out_punch: bool
    unconfirmed_in: bool
    unconfirmed_out: bool
    late_out_to_meal: bool

    def __str__(self) -> str:
        out = []
        for name, value in self.to_dict().items():
            if value:
                if isinstance(value, bool):
                    out.append(name)
                else:
                    out.append(f'{name}={value}')
        return '\n'.join(out)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "lOvertimeException": self.overtime,
            "lTardyException": self.tardy,
            "lEarlyOutException": self.early_out,
            "lEarlyInException": self.early_in,
            "lLateOutException": self.late_out,
            "lMissingPunchException": self.missing_punch,
            "lMealBreakPremiumException": self.meal_break_premium,
            "cMealException": self.meal_desc,
            "lUnderPayException": self.under_pay,
            "lOverPayException": self.over_pay,
            "lTardyGraceException": self.tardy_grace,
            "lEarlyOutGraceException": self.early_out_grace,
            "lUnpaidBreakException": self.unpaid_break,
            "cBreakException": self.break_desc,
            "lAutoDeductMealException": self.auto_deduct_meal,
            "lAutoMealWaivedException": self.auto_meal_waived,
            "lUnconfirmedPunchException": self.unconfirmed_punch,
            "lUnconfirmedInPunch": self.unconfirmed_in_punch,
            "lUnconfirmedOutPunch": self.unconfirmed_out_punch,
            "lUnconfirmedInException": self.unconfirmed_in,
            "lUnconfirmedOutException": self.unconfirmed_out,
            "lLateOutToMealException": self.late_out_to_meal,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        yield f"[b]TimesheetEntryException:[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("overtime", str(self.overtime))
        my_table.add_row("tardy", str(self.tardy))
        my_table.add_row("early_out", str(self.early_out))
        my_table.add_row("early_in", str(self.early_in))
        my_table.add_row("late_out", str(self.late_out))
        my_table.add_row("missing_punch", str(self.missing_punch))
        my_table.add_row("meal_break_premium", str(self.meal_break_premium))
        my_table.add_row("meal_desc", self.meal_desc)
        my_table.add_row("under_pay", str(self.under_pay))
        my_table.add_row("over_pay", str(self.over_pay))
        my_table.add_row("tardy_grace", str(self.tardy_grace))
        my_table.add_row("early_out_grace", str(self.early_out_grace))
        my_table.add_row("unpaid_break", str(self.unpaid_break))
        my_table.add_row("break_desc", self.break_desc)
        my_table.add_row("auto_deduct_meal", str(self.auto_deduct_meal))
        my_table.add_row("auto_meal_waived", str(self.auto_meal_waived))
        my_table.add_row("unconfirmed_punch", str(self.unconfirmed_punch))
        my_table.add_row("unconfirmed_in_punch", str(
            self.unconfirmed_in_punch))
        my_table.add_row("unconfirmed_out_punch", str(
            self.unconfirmed_out_punch))
        my_table.add_row("unconfirmed_in", str(self.unconfirmed_in))
        my_table.add_row("unconfirmed_out", str(self.unconfirmed_out))
        my_table.add_row("late_out_to_meal", str(self.late_out_to_meal))
        yield my_table


@dataclass
class Punch:
    punch: arrow.arrow.Arrow
    adjust: arrow.arrow.Arrow
    modified: bool
    gps: GPS
    site: str
    og: arrow.arrow.Arrow
    net_chk_fail: bool
    expression: str
    expression_save: str
    timezone: arrow.arrow.Arrow
    recording: Any


@dataclass
class TimesheetEntry:
    """Representation of a time entry in a NOVATime timesheet."""
    # pylint: disable=too-many-instance-attributes
    # There is no good way to package these into fewer attributes, so pylint can blow it
    sheet_sequence: int     # iTimesheetSeq
    pay_period: DatePeriod   # dPayPeriodStart, dPayPeriodEnd
    work_period: DatePeriod   # dWorkPeriodStartDate, dWorkPeriodEndDate
    entry_sequence: int     # iTimeSeq

    # iEmployeeSeq, cEmployeeID, cEmployeeFirstName, cEmployeeLastName, cEmployeeFullName
    employee: User

    punch_date: arrow.arrow.Arrow   # dPunchDate (probably dWorkDate?)
    work_date: arrow.arrow.Arrow    # dWorkDate

    # tPunchDateTime (probably punch_in.start?)
    punch_date_time: arrow.arrow.Arrow

    # dIn, nAdjustIn, lInMod, cInGPS, cSiteIn, dOGIn, lInNetChkFail, cInExpression, cInExpressionSave, nTZIn, mInRecording
    punch_in: Punch

    # dOut, nAdjustOut, lOutMod, cOutGPS, cSiteOut, dOGOut, lOutNetChkFail, cOutExpression, cOutExpressionSave, nTZOut, mOutRecording
    punch_out: Punch

    pay_code: PayCode       # nPayCode, cExpCode, cPayCodeDescription, nCodeType
    categories: List[EntryCategory]
    exceptions: List[TimesheetEntryException]
    shift_expression: str

    def __init__(self, entry: Dict[str, Any]) -> None:
        pass

    def work_hours(self) -> float:
        pass

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "iTimesheetSeq": 0,
            "dPayPeriodStart": self.timesheet.pay_period.start.format(PUNCH_FORMAT),
            "dPayPeriodEnd": self.timesheet.pay_period.end.format(PUNCH_FORMAT),
            "iTimeSeq": self.sequence,
            "iEmployeeSeq": self.timesheet.novatime.user.employee_seq,
            "cEmployeeID": str(self.timesheet.novatime.user.username),
            "cEmployeeFirstName": self.timesheet.novatime.user.first_name,
            "cEmployeeLastName": self.timesheet.novatime.user.last_name,
            "cEmployeeFullName": self.timesheet.novatime.user.full_name,
            "dPunchDate": self.punch_date.format(PUNCH_FORMAT),
            "dWorkDate": self.work_date.format(PUNCH_FORMAT),
            "nPayCode": self.pay_code.code,
            "cPayCodeDescription": str(self.pay_code),
            "nCodeType": self.pay_code.code_type,
            "nCalculate": 1,
            "lComputeNonCalc": False,
            "cExpCode": self.pay_code.description,
            "dIn": self.punch_in.format(PUNCH_FORMAT),
            "dOut": self.punch_out.format(PUNCH_FORMAT),
            "dOGIn": self.punch_in.format(PUNCH_FORMAT),
            "dOGOut": self.punch_out.format(PUNCH_FORMAT),
            "nScheduleHours": 0.0,
            "nWorkHours": self.work_hours(),
            "nOT1Hours": 0.0,
            "nOT2Hours": 0.0,
            "nOT3Hours": 0.0,
            "nOT4Hours": 0.0,
            "nOT5Hours": 0.0,
            "nPayAmount": 0.0,
            "OTTotalHoursOnePunch": 0.0,
            "nPayRate": 0.0,
            "lCalcOverride": False,
            "iNoteSeq": 0,
            "cAuthor": "",
            "cNotes": "",
            "cReasonCode": "",
            "cReasonColor": "",
            "nApprovalStatus": 0,
            "lApprovalStatus": False,
            "GroupValueList": [category.to_dict() for category in self.categories],
            "AccessibleGroupList": None,
            "nQuantityGood": 0.0,
            "nQuantityBad": 0.0,
            "cInGPS": None,
            "cOutGPS": "",
            "cMoreInfo": "",
            "lPendingCalc": False,
            "cShiftExpression": self.shift_expression,
            "cSiteIn": self.site_in,
            "cSiteOut": self.site_out,
            "lElapsedTime": False,
            "lInMod": False,
            "lOutMod": False,
            "lInNetChkFail": False,
            "lOutNetChkFail": False,
            "lUnauthorizedOT": False,
            "cInExpression": "    ",
            "cOutExpression": "    ",
            "cInExpressionSave": "    ",
            "cOutExpressionSave": "    ",
            "cGroupCode": " ",
            "cSchedule": "",
            "cInOutExpression": "NM1",
            "nCompOT1Hours": 0.0,
            "nCompOT2Hours": 0.0,
            "nCompOT3Hours": 0.0,
            "nCompOT4Hours": 0.0,
            "nCompOT5Hours": 0.0,
            "nRawCompOT1Hours": 0.0,
            "nRawCompOT2Hours": 0.0,
            "nRawCompOT3Hours": 0.0,
            "nRawCompOT4Hours": 0.0,
            "nRawCompOT5Hours": 0.0,
            "nRedirectOT1Hours": 0.0,
            "nRedirectOT2Hours": 0.0,
            "nRedirectOT3Hours": 0.0,
            "nRedirectOT4Hours": 0.0,
            "nRedirectOT5Hours": 0.0,
            "nAdjustIn": self.punch_in.format(PUNCH_TIME_FORMAT),
            "nAdjustOut": self.punch_out.format(PUNCH_TIME_FORMAT),
            "GroupingString": ','.join([category.group_seq for category in self.categories]),
            "SchGroupingString": None,
            "RecordType": 1,
            "lShowSchedulePayCode": False,
            "lReadOnly": False,
            "lPayCodeReadOnly": False,
            "dAdjustmentDate": None,
            "lReversed": False,
            "dReverseDate": None,
            "nReverseStatus": 0,
            "WeekGroupString": "06/20/2022 - 06/26/2022",
            "Grouping": {
                "iGroupNumber": 6,
                "iGroupValueSeq": 1810,
                "cGroupValue": "1",
                "cGroupValueDescription": "Unassigned",
                "cGroupCode": None,
                "iGroupUserType": 0,
                "nGPSLatitude": None,
                "nGPSLongitude": None,
                "cGroupColor": None,
                "cDescription": None,
                "lClosed": None
            },
            "cPayType": " ",
            "nRegPay": 0.0,
            "nOT1Pay": 0.0,
            "nOT2Pay": 0.0,
            "nOT3Pay": 0.0,
            "nOT4Pay": 0.0,
            "nOT5Pay": 0.0,
            "nPremiumPay": 0.0,
            "nTotalPay": 0.000,
            "cGroupLevel": "800",
            "dWorkPeriodStartDate": None,
            "dWorkPeriodEndDate": None,
            "tLastModified": None,
            "lAudit": True,
            "lRefTime": False,
            "lIsTGARecord": False,
            "nTZIn": None,
            "nTZOut": None,
            "lPending": False,
            "DateKey": self.punch_date.format(PUNCH_DATEKEY_FORMAT),
            "lWithinPP": True,
            "tPunchDateTime": self.punch_in.format(PUNCH_FORMAT),
            "CarryoverExpansionOverride": 0,
            "lCarryoverExpansionORChanged": False,
            "ExpectedMealTimes": [
                {
                    "iIndex": 1,
                    "tStartTime": "06/20/2022 14:41:00",
                    "tEndTime": "06/20/2022 15:11:00"
                }
            ],
            "InvalidGroupList": [],
            "nLongMeal": 0.0,
            "lHasLstChgDay": False,
            "nTardyMinutes": 0,
            "nEarlyOutMinutes": 0,
            "nAutoMealMinutes": 30.0,
            "nMealValMinutes": 0.0,
            "mInRecording": None,
            "mOutRecording": None,
            "lAutoPayNoDelete": False,
            "lNonComputeCalcPayCode": False,
            "cAssignID": "",
            "nPayPolicy": 1.0,
            "cPayPolicyDescription": "1[HOURLY DAYS TIME CLOCK]",
            "ReadOnlyReasonCodeDescription": None,
            "isSchedulePremium": False,
            "isSchedulePremiumUserOverride": False
        }
        for exception in self.exceptions:
            d.update(exception.to_dict())


class NOVATime:
    """Wrapper for the "REST API" for NOVATime"""
    timesheet: 'Timesheet'
    groups: Dict[str, EntryCategoryGroup]

    def __init__(self):
        self.logger = logging.getLogger('NOVATime')
        # set up secrets info
        if not isfile('secrets.ini'):
            raise FileNotFoundError(
                'Please copy secrets.ini.example to secrets.ini and configure per the comments')

        self.logger.debug('Loading secrets file')
        self._secrets = configparser.ConfigParser()
        self._secrets.optionxform = lambda option: option  # return case-sensitive keys
        self._secrets.read('secrets.ini')
        self.cid = self._secrets['uri']['cid']
        self.host = self._secrets['uri']['host']
        self.page = self._secrets['uri']['page']
        self.api_url = f'{self.host}/{self.page}/{self.cid}'
        self.logger.debug(f'API URI: {self.api_url}')

        self.user = User(
            username=self._secrets['user']['user'],
            password=self._secrets['user']['password']
        )

        self.groups = {}

        self._session = requests.Session()

        # set up headers in secrets and session
        headers = {
            'accept':   'text/html,'
                        'application/xhtml+xml,'
                        'application/xml;q=0.9,'
                        'image/avif,'
                        'image/webp,'
                        '*/*;q=0.8',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'en-US,en;q=0.5',
            'connection': 'keep-alive',
            'content-type': 'application/x-www-form-urlencoded',
            'dnt': '1',
            'locale': 'en-US',
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:68.0) Gecko/20100101 Firefox/68.0',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-user': '?1',
            'sec-fetch-site': 'same-origin',
        }
        self._secrets.read_dict({'header': headers})
        self._session.headers.update(headers)

    def _ado_to_dict(self, datalist):
        """
        Clean up the ADO DataList objects that sometimes come from NOVATime

        Args:
            datalist (dict): the ADO DataList dict, with Key and Value keys

        Returns:
            dict[str, str]: the resulting Pythonic dict
        """
        self.logger.debug(f'ADO: {datalist}')
        if 'DataList' in datalist:
            datalist = datalist['DataList']
        data = {}
        for obj in datalist:
            data[obj['Key']] = obj['Value']
        return data

    def _build_login_data(self, loginpage):
        """
        Build a dict for NOVATime login request

        Args:
            loginpage (BeautifulSoup): a bs4 parser of the login page HTML

        Returns:
            dict[str,str]: dict for login request
        """
        data = {
            '__EVENTTARGET': loginpage.find(id='__EVENTTARGET')['value'],
            '__EVENTARGUMENT': loginpage.find(id='__EVENTARGUMENT')['value'],
            '__VIEWSTATE': loginpage.find(id='__VIEWSTATE')['value'],
            '__VIEWSTATEGENERATOR': loginpage.find(id='__VIEWSTATEGENERATOR')['value'],
            '__VIEWSTATEENCRYPTED': loginpage.find(id='__VIEWSTATEENCRYPTED')['value'],
            '__RequestVerificationToken': loginpage.find(attrs={'name': '__RequestVerificationToken'})['value'],
            'txtUserName': self.user.username,
            'txtPassword': self.user.password,
            "hUserAgent": self._secrets['header']['user-agent'],
            "changePWsecq1$txtOldPW": "",
            "changePWsecq1$txtPassword": "",
            "changePWsecq1$txtPWVerify": "",
            "changePWsecq1$SecqDdl": "0",
            "changePWsecq1$txtAnswer": "",
            "multiFactorAuth$codeEntryTxt": "",
            "multiFactorAuth$hdnFldcurrSeq": "",
            "multiFactorAuth$hdnFldconsumerType": "",
            "txtPunchMsg": "",
            "btnLogin": "Employee+Web+Services",
            "hiddenGPSCoords": "",
            "btnLogin_PosX": "118",
            "btnLogin_PosY": "450",
            "hCpuClass": "undefined",
            "hBrowserName": "Netscape",
            "hBrowserVersion": "5.0+(Windows)",
            "hUserPlatform": "Win32",
            "hScreenWidth": "1920",
            "hScreenHeight": "1080",
            "totalCount": "0"
        }

        for hidden_input in loginpage.find_all(name='input', type='hidden', value=True):
            if hidden_input.contents:
                continue
            id = ''
            name = ''
            value = ''
            if hidden_input.has_attr('id'):
                id = hidden_input['id']
            if hidden_input.has_attr('name'):
                name = hidden_input['name']
            if hidden_input.has_attr('value'):
                value = hidden_input['value']
            id = id or name
            if id not in data:
                self.logger.warning(
                    f'New input field not in request header: id={id} value={value}')

        return data

    def login(self):
        """Login to NOVATime"""

        # build uri
        uri = f"https://{self.host}/novatime/ewskiosk.aspx"
        self.logger.debug(
            f'Requesting NOVATime landing page {uri}')

        # get the login landing page to harvest for login request data
        loginrequest = self._session.get(uri, params={'CID': self.cid})
        loginpage = BeautifulSoup(loginrequest.text, 'html.parser')

        # update secrets and session with login cookie parameters
        self._secrets.read_dict(
            {'cookie': loginrequest.cookies.get_dict(domain=self.host, path='/')})

        for cookie, value in self._secrets['cookie'].items():
            self._session.cookies.set(
                cookie, value, domain=self.host, path='/')

        # POST the login request to the server
        login_data = self._build_login_data(loginpage=loginpage)
        self.logger.debug(
            f'Logging into NOVATime with user {self.user.username} details: {login_data}')
        response = self._session.post(uri, params={'CID': self.cid},
                                      data=login_data)
        if not response.ok:
            raise ValueError(
                f'Bad response: {response.status_code} - {response.reason}')

        # the SessionVariable contains several pieces of information we need for future requests;
        #  add them to the session headers and the secrets
        self.logger.debug(
            f'Asking for NOVATime user details for {self.user.username}')
        user_data_request = self._session.get(
            url=f'https://{self.api_url}/SessionVariable')

        user_data = self._ado_to_dict(user_data_request.json())
        if user_data['USERSEQ']:
            self.user.user_seq = user_data['USERSEQ']
        else:
            self.user.user_seq = '0'
        self.user.employee_seq = user_data['EMPSEQ']

        pay_period = DatePeriod(
            start=parse_date(user_data['PPSTART']),
            end=parse_date(user_data['PPEND'])
        )

        self.logger.debug(f'User {self.user.username} has '
                          f'user_seq {self.user.user_seq} and '
                          f'employee_seq {self.user.employee_seq}, '
                          f'with pay period {pay_period}')

        self.timesheet = Timesheet(novatime=self,
                                   pay_period=pay_period,
                                   weekhours=timedelta(
                                       hours=int(self._secrets['hours']['weekhours']))
                                   )

        self._session.headers.update(
            {'EmployeeSeq': self.user.employee_seq,
             'UserSeq': self.user.user_seq})

        # fill out the user record with other data
        self.logger.debug(
            f'Asking for NOVATime employee record for {self.user.username}')
        user_data = self._session.get(
            url=f'https://{self.api_url}/employee/{self.user.employee_seq}').json()

        self.user.access_seq = str(user_data['Data']['iAccessSeq'])
        self.user.first_name = user_data['Data']['cFirstName']
        self.user.last_name = user_data['Data']['cLastName']
        self.user.full_name = user_data['Data']['cFullName']

        self.logger.debug(f'User {self.user.username} is '
                          f'{self.user.first_name} {self.user.last_name} '
                          f'({self.user.full_name}) with access_seq {self.user.access_seq}')
        self.get_groups()
        for group in self.groups:
            self.get_group_options(group=group)

    def get_timesheet(self, pay_period: DatePeriod = None):
        """Download the timesheet and populate internal data.

        Here there be dragons. This does no error-checking, and is brittle as most scrapers are.

        Args:
            pay_period (PayPeriod): the pay period for the timesheet
                                    (default None, for login pay period)

        """
        if pay_period is not None:
            timesheet = Timesheet(novatime=self,
                                  pay_period=pay_period,
                                  weekhours=timedelta(
                                      hours=int(self._secrets['hours']['weekhours']))
                                  )
        else:
            timesheet = self.timesheet
            pay_period = timesheet.pay_period

        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/timesheetdetail'
        parameters = {
            'AccessSeq': self.user.access_seq,
            'EmployeeSeq': self.user.employee_seq,
            'UserSeq': self.user.user_seq,
            'StartDate': pay_period.start.format(PARAM_DATE_FORMAT),
            'EndDate': pay_period.end.format(PARAM_DATE_FORMAT),
            'CustomDateRange': False,
            'ShowOneMoreDay': False,
            'EmployeeSeqList': '',
            'DailyDate': pay_period.start.format(PARAM_DATE_FORMAT),
            'ForceAbsent': False,
            'PolicyGroup': ''
        }

        response = self._session.get(uri, params=parameters)
        raw_timesheet = response.json()

        # if the user is not authed, this pukes so handle it gracefully-ish <3
        if 'DataList' not in raw_timesheet:
            raise AttributeError('Not authorized, please check secrets.ini')

        # grab current pay period timesheet, write to JSON file in safe directory
        if not isdir('pay'):
            os.mkdir('pay')
        with open(join('pay', f'{pay_period}.json'),
                  encoding='utf-8',
                  mode='w') as times:
            json.dump(raw_timesheet, times, indent=' '*4)
        return raw_timesheet
        raw_timesheet = raw_timesheet['DataList']

        # ok we made it! process the current pay period
        self.timesheet.parse_timesheet(raw_timesheet)

    def get_groups(self) -> Dict[str, EntryCategoryGroup]:
        """
        Fetch the mapping of groups

        Returns:
            Dict[str,EntryCategoryGroup]: the internal dict mapping group name to details
        """
        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/systemsetting'
        self.logger.debug('Getting groups:')
        response = self._session.get(uri)
        for group in response.json()['GROUPLIST']:
            self.groups[group['cGroupCaption']] = EntryCategoryGroup(
                value=group['iGroupNumber'],
                name=group['cGroupCaption'],
                options={}
            )
            self.logger.debug(
                f"\t{group['cGroupCaption']}: {group['iGroupNumber']}")
        return self.groups

    def get_group_options(self, group: str) -> EntryCategoryGroup:
        """
        Fetch the options for a given group.

        Args:
            group (str): One of the groups in self.group

        Raises:
            ValueError: on error fetching group options

        Returns:
            EntryCategoryGroup: the options for the given group
        """
        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/Group/GetPagedGroups'
        items_per_page = 10
        page = 1
        group_options = self.groups[group]
        group_options.options = {}
        parameters = {
            'GroupNumber': str(group_options.value),
            'UserAccessSeq': str(self.user.user_seq),
            'EmployeeAccessSeq': str(self.user.access_seq),
            'PrimaryGroupNumber': '0',
            'PrimaryGroupValueSeq': '0',
            'PrimaryGroup2Number': '0',
            'PrimaryGroup2ValueSeq': '0',
            'PrimaryGroup3Number': '0',
            'PrimaryGroup3ValueSeq': '0',
            'PrimaryGroup4Number': '0',
            'PrimaryGroup4ValueSeq': '0',
            'UserSeq': str(self.user.user_seq),
            'EmployeeSeq': str(self.user.employee_seq),
            'UseCascadingGroupLinkage': 'false',
            'CurrentPage': str(page),
            'ItemsPerPage': str(items_per_page),
            'SearchText': '',
            'SelectedEmployeeSeq': str(self.user.employee_seq),
            'FilterJobGroups': 'false',
        }

        self.logger.debug(f'Getting options for group {group_options.name}')
        with Progress(transient=True) as progress:
            group_task = progress.add_task(
                f'Getting options for group {group_options.name}:', total=None)
            while not progress.finished:
                response = self._session.get(uri, params=parameters)
                group_data = response.json()
                if group_data['_errorCode'] != 1:
                    raise ValueError(
                        f"Error getting group {group_options.name}: {group_data['_errorCode']} - "
                        f"{group_data['_errorDescription']}")
                items = group_data['Data']['ItemTotal']
                if page == 1:
                    progress.update(group_task, total=items)
                    progress.start_task(group_task)
                progress.update(group_task, advance=len(
                    group_data['Data']['PagedList']))
                for option in map(EntryCategory, group_data['Data']['PagedList']):
                    if option.group_seq in group_options.options:
                        old = group_options.options[option.group_seq]
                        new = option
                        self.logger.warning(f'Overwriting {option.group_seq}:'
                                            f'{old} ({old.group_seq})'
                                            f'->{new} ({new.group_seq})')
                    group_options.options[option.group_seq] = option
                if len(group_options.options) < items:
                    page += 1
                    parameters['CurrentPage'] = str(page)
                    continue
                self.logger.debug(
                    f'Got {len(group_options.options)} of {items} for {group_options.name}')
            self.groups[group_options.name] = group_options
        return group_options


class Timesheet:
    """Represent a NOVATime timesheet for a pay period"""
    entries = Dict[arrow.arrow.Arrow, TimesheetEntry]
    hours: HourTotals
    exceptions = None
    pay_period: DatePeriod
    remaining: timedelta

    def __init__(self, novatime: NOVATime, pay_period: DatePeriod, weekhours: timedelta):
        self.novatime = novatime
        self.pay_period = pay_period
        self.weekhours = weekhours

    def format_td(self, delta: timedelta) -> str:
        """Format a `datetime.timedelta` as hours and minutes"""
        hours, remainder = divmod(delta.total_seconds(), 3600)
        minutes, _ = divmod(remainder, 60)
        return f'{hours:.0f}:{minutes:02.0f}'

    def parse_timesheet(self, raw_timesheet):
        self.get_times(raw_timesheet)
        self.get_exceptions()

    def make_timesheet_report(self):
        """Print a little timesheet report."""

        print(f'Pay period from {self.pay_period.start.date()}'
              f' to {self.pay_period.end.date()}:')
        if self.exceptions:
            print('\tExceptions:')
            for date, codes in self.exceptions.items():
                print(f'\t\t{date.format("dddd, MMMM DD, YYYY")}:')
                for code, value in codes.items():
                    print(f'\t\t\t{code} = {value}')
        if self.hours.last_week:
            print(f'\tLast week: {self.format_td(self.hours.last_week)}')

        # then, more usefully, report how much time left this week,
        #  and, if there is a missing punch today (AKA we are clocked in!), report when to clock out
        #
        # (this if course does no bounds checking or anything so it probably does
        #  entertaining things in overtime or Tuesday conditions)
        self.remaining = self.weekhours - self.hours.this_week
        print(
            f'\tThis week: {self.format_td(self.hours.this_week)} ({self.format_td(self.remaining)} left)')

        clock_in, clock_out = self.predict_clock_out(self.remaining)

        if clock_out is not None:
            print(f'\tAfter clocking in at {clock_in.format("HH:mm")},'
                  f' clock out by {clock_out.format("HH:mm")}'
                  f' to hit {self.format_td(self.weekhours)}')

    def get_exceptions(self):
        """Retrieve any exceptions during the pay period.

        There are likely lots of them, and many are innocuous.
        """
        self.exceptions = {}
        for entry in self.entries:
            punch_date = parse_punch(entry['dPunchDate'])
            self.exceptions[punch_date] = {key: entry[key]
                                           for key in entry if 'Exception' in key and entry[key]}
            if not self.exceptions[punch_date]:
                del self.exceptions[punch_date]

    def get_times(self, raw_timesheet):
        """Retrieve the daily hours during the pay period."""
        self.entries = {}
        self.hours.last_week = timedelta(hours=0)
        self.hours.this_week = timedelta(hours=0)
        self.hours.total = timedelta(hours=0)
        for entry in raw_timesheet:
            punch_date = parse_punch(entry['dPunchDate'])
            self.entries[punch_date] = timedelta(hours=entry['nDailyHours'])
            if self.is_this_week(punch_date):
                self.hours.this_week += self.entries[punch_date]
            else:
                self.hours.last_week += self.entries[punch_date]
            self.hours.total += self.entries[punch_date]

    def is_this_week(self, date):
        """Report whether a date is in this week"""
        this_sunday = arrow.now().shift(weekday=6).floor('day')
        last_sunday = this_sunday.shift(weeks=-1).floor('day')
        return date.is_between(last_sunday, this_sunday, '[)')

    def predict_clock_out(self, remaining):
        """Try to predict an appropriate clock-out time.

        Args:
            remaining (timedelta): the remaining hours this week

        Returns:
            tuple[arrow, arrow]: the first is an `arrow` if the day has a clock-in time, or None if not
                                the second is an `arrow` of clock_in plus remaining plus lunch
                                    if remaining is longer than 8 hours; however, if there is no
                                    "missing punch" exception (AKA we are not clocked in!),
                                    this will be None

        """
        today = arrow.now().floor('day')
        for entry in self.entries:
            if parse_punch(entry['dPunchDate']) == today:
                today = entry
                # found today's entry
                break

        if isinstance(today, arrow.arrow.Arrow):
            # no entry today
            return None, None
        clock_in = parse_punch(today['dIn'])
        if today['lMissingPunchException']:
            # we are clocked in, figure out a good clock-out time
            clock_out = clock_in + remaining
            if remaining > timedelta(hours=8):
                # long enough shift we need a lunch
                clock_out = clock_out.shift(minutes=today['nAutoMealMinutes'])
        else:
            # not clocked in
            clock_out = None
        return clock_in, clock_out


if __name__ == '__main__':
    n = NOVATime()
    n.login()
