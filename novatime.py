# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL

"""Wrapper for the "REST API" for NOVATime"""

import configparser
import json
import logging
import os
import sys
from datetime import timedelta
from os.path import isdir, isfile, join
from typing import Any, Dict, List, NamedTuple

import arrow
import requests
from bs4 import BeautifulSoup
from rich.logging import RichHandler
from rich.progress import Progress

PARAM_DATE_FORMAT = 'ddd MMM DD YYYY'
PUNCH_FORMAT = 'MM/DD/YYYY HH:mm:ss'
PUNCH_TIME_FORMAT = 'HH:mmA'
NOVA_DATE_FORMAT = 'M/D/YYYY'
PUNCH_DATEKEY_FORMAT = 'MM/DD'


FORMAT = "%(message)s"
logging.basicConfig(
    level="DEBUG", format=FORMAT, datefmt="[%X]", handlers=[RichHandler(rich_tracebacks=True, tracebacks_suppress=[requests])]
)
logging.getLogger("requests").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)


def parse_punch(punch_str):
    """Parse a punch entry into an `arrow`"""
    return arrow.get(punch_str, PUNCH_FORMAT).replace(tzinfo='America/Detroit')


def parse_date(date_str):
    """Parse a date entry into an `arrow`"""
    return arrow.get(date_str, NOVA_DATE_FORMAT).replace(tzinfo='America/Detroit')


class PayPeriod:
    start: arrow.arrow.Arrow
    end: arrow.arrow.Arrow

    def __str__(self) -> str:
        return f'{self.start.date()}--{self.end.date()}'


class User:
    username: str
    password: str
    user_seq: str
    employee_seq: str
    access_seq: str
    first_name: str
    last_name: str
    full_name: str


class HourTotals:
    last_week: timedelta
    this_week: timedelta
    total: timedelta


class EntryCategory:
    value: int
    description: str
    group_number: int
    group_seq: int
    group_code: str
    group_user_type: int
    gps_latitude: float
    gps_longitude: float
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
        self.gps_latitude = group['nGPSLatitude']
        self.gps_longitude = group['nGPSLongitude']

    def __str__(self) -> str:
        return f'{self.value} [{self.description or self.group_value_description}]'

    def to_dict(self) -> Dict[str, Any]:
        return {
            "iGroupNumber": self.group_number,
            "iGroupValueSeq": self.group_seq,
            "cGroupValue": str(self.value),
            "cGroupValueDescription": self.group_value_description,
            "cGroupCode": self.group_code,
            "iGroupUserType": self.group_user_type,
            "nGPSLatitude": self.gps_latitude,
            "nGPSLongitude": self.gps_longitude,
            "cGroupColor": self.group_color,
            "cDescription": self.description,
            "lClosed": self.closed
        }


class PayCode:
    code: int
    description: str
    code_type: int

    def __str__(self) -> str:
        return f'{self.code}[{self.description}]'


class TimesheetEntryException:
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


class TimesheetEntry:
    timesheet: 'Timesheet'
    sequence: int
    pay_code: PayCode
    punch_date: arrow.arrow.Arrow
    work_date: arrow.arrow.Arrow
    punch_in: arrow.arrow.Arrow
    punch_out: arrow.arrow.Arrow
    categories: List[EntryCategory]
    exceptions: List[TimesheetEntryException]
    shift_expression: str
    site_in: str
    site_out: str

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
    user: User

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

        self.user = User()
        self.user.username = self._secrets['user']['user']
        self.user.password = self._secrets['user']['password']

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

        pay_period = PayPeriod()
        pay_period.start = parse_date(user_data['PPSTART'])
        pay_period.end = parse_date(user_data['PPEND'])

        self.logger.debug(
            f'User {self.user.username} has user_seq {self.user.user_seq} and employee_seq {self.user.employee_seq}, with pay period {pay_period}')

        self.timesheet = Timesheet(novatime=self,
                                   pay_period=pay_period,
                                   weekhours=timedelta(
                                       hours=int(self._secrets['hours']['weekhours']))
                                   )

        self._session.headers.update(
            {'EmployeeSeq': self.user.employee_seq,
             'UserSeq': self.user.user_seq})

        # the employee record contains an access ID we need for other requests, add it to the secrets
        self.logger.debug(
            f'Asking for NOVATime employee record for {self.user.username}')
        user_data = self._session.get(
            url=f'https://{self.api_url}/employee/{self.user.employee_seq}').json()

        self.user.access_seq = str(user_data['Data']['iAccessSeq'])
        self.user.first_name = user_data['Data']['cFirstName']
        self.user.last_name = user_data['Data']['cLastName']
        self.user.full_name = user_data['Data']['cFullName']

        self.logger.debug(
            f'User {self.user.username} is {self.user.first_name} {self.user.last_name} ({self.user.full_name}) with access_seq {self.user.access_seq}')
        self.get_groups()

    def get_timesheet(self, pay_period: PayPeriod = None):
        """Download the timesheet and populate internal data.

        Here there be dragons. This does no error-checking, and is brittle as most scrapers are.

        Args:
            pay_period (PayPeriod): the pay period for the timesheet
                                    (default None, for login pay period)

        """
        if pay_period is not None:
            timesheet = Timesheet(novatime=self,
                                  pay_period=pay_period,
                                  weekhours=timedelta(hours=int(self._secrets['hours']['weekhours'])))
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

    def get_groups(self):
        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/systemsetting'
        self.logger.debug('Getting groups:')
        response = self._session.get(uri)
        self.groups = {}
        for group in response.json()['GROUPLIST']:
            self.groups[group['cGroupCaption']] = group['iGroupNumber']
            self.logger.debug(
                f"\t{group['cGroupCaption']}: {group['iGroupNumber']}")
        return self.groups

    def get_group_options(self, group):
        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/Group/GetPagedGroups'
        items_per_page = 10
        page = 1
        parameters = {
            'GroupNumber': str(self.groups[group]),
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
        data = []
        self.logger.debug(f'Getting options for group {group}')
        with Progress(transient=True) as progress:
            group_task = progress.add_task(
                f'Getting options for group {group}:', total=None)
            while not progress.finished:
                response = self._session.get(uri, params=parameters)
                group_data = response.json()
                if group_data['_errorCode'] != 1:
                    raise ValueError(
                        f"Error getting group {group}: {group_data['_errorCode']} - {group_data['_errorDescription']}")
                items = group_data['Data']['ItemTotal']
                if page == 1:
                    progress.update(group_task, total=items)
                    progress.start_task(group_task)
                progress.update(group_task, advance=len(
                    group_data['Data']['PagedList']))
                data += group_data['Data']['PagedList']
                if len(data) < items:
                    page += 1
                    parameters['CurrentPage'] = str(page)
                    continue
                else:
                    self.logger.debug(f'Got {len(data)} of {items} for {group}')
        return list(map(EntryCategory, data))


class Timesheet:
    """Represent a NOVATime timesheet for a pay period"""
    entries = Dict[arrow.arrow.Arrow, TimesheetEntry]
    hours: HourTotals
    exceptions = None
    pay_period: PayPeriod
    remaining: timedelta

    def __init__(self, novatime: NOVATime, pay_period: PayPeriod, weekhours: timedelta):
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
