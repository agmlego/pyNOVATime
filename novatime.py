# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL
# pylint: disable=logging-fstring-interpolation

"""Wrapper for the "REST API" for NOVATime"""

import configparser
import json
import logging
import os
import shelve
from dataclasses import dataclass
from datetime import datetime, timedelta
from os.path import isdir, isfile, join
from typing import Any, Dict, List, Tuple, Union

import arrow
import requests
# HACK Until BRSU fixes the SSL proxy, this makes things quieter with `verify=False`
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
# end HACK
from bs4 import BeautifulSoup
from rich import print
from rich.console import Console, ConsoleOptions, RenderResult
from rich.logging import RichHandler
from rich.progress import Progress
from rich.table import Table



PARAM_DATE_FORMAT = 'ddd MMM DD YYYY'
PUNCH_FORMAT = 'MM/DD/YYYY HH:mm:ss'
PUNCH_TIME_FORMAT = 'HH:mmA'
NOVA_DATE_FORMAT = 'M/D/YYYY'
PUNCH_DATEKEY_FORMAT = 'MM/DD'
WRITE_ENTRY_FORMAT = 'YYYY-MM-DDTHH:mm:ss.SSS[Z]'


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


def parse_punch(punch_str: Union[str, arrow.arrow.Arrow]):
    """Parse a punch entry into an `arrow`"""
    if isinstance(punch_str, str):
        return arrow.get(punch_str, PUNCH_FORMAT).replace(tzinfo='America/Detroit')
    return punch_str


def parse_date(date_str):
    """Parse a date entry into an `arrow`"""
    return arrow.get(date_str, NOVA_DATE_FORMAT).replace(tzinfo='America/Detroit')


@dataclass
class DatePeriod:
    """Pay period with start and end dates."""
    start: arrow.arrow.Arrow
    end: arrow.arrow.Arrow

    def __init__(self, start: Union[str, arrow.arrow.Arrow], end: Union[str, arrow.arrow.Arrow]) -> None:
        self.start = parse_punch(start)
        self.end = parse_punch(end)

    def __str__(self) -> str:
        return f'{self.start.date()}--{self.end.date()}'

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]DatePeriod:[/b] {self.start} to {self.end}"


@dataclass
class GPS:
    """GPS coordinates."""
    latitude: float
    longitude: float

    def __str__(self) -> str:
        return f'{self.latitude}, {self.longitude}'

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]GPS:[/b] {self.latitude}, {self.longitude}"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "nGPSLatitude": self.latitude,
            "nGPSLongitude": self.longitude,
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
        # pylint: disable=unused-argument
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
class HoursGroup:
    """Representation of OT hours fields in a NOVATime timesheet."""
    overtime_1: timedelta  # n*OT1Hours
    overtime_2: timedelta  # n*OT2Hours
    overtime_3: timedelta  # n*OT3Hours
    overtime_4: timedelta  # n*OT4Hours
    overtime_5: timedelta  # n*OT5Hours
    ot_type: str           # *

    def __init__(self,
                 overtime_1: float,
                 overtime_2: float,
                 overtime_3: float,
                 overtime_4: float,
                 overtime_5: float,
                 ot_type: str
                 ) -> None:
        self.overtime_1 = timedelta(hours=overtime_1)
        self.overtime_2 = timedelta(hours=overtime_2)
        self.overtime_3 = timedelta(hours=overtime_3)
        self.overtime_4 = timedelta(hours=overtime_4)
        self.overtime_5 = timedelta(hours=overtime_5)
        self.ot_type = ot_type

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
            f"n{self.ot_type}OT1Hours": self.overtime_1.total_seconds()/3600.0,
            f"n{self.ot_type}OT2Hours": self.overtime_2.total_seconds()/3600.0,
            f"n{self.ot_type}OT3Hours": self.overtime_3.total_seconds()/3600.0,
            f"n{self.ot_type}OT4Hours": self.overtime_4.total_seconds()/3600.0,
            f"n{self.ot_type}OT5Hours": self.overtime_5.total_seconds()/3600.0,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]HourTotals: {self.ot_type}[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("overtime_1", str(self.overtime_1))
        my_table.add_row("overtime_2", str(self.overtime_2))
        my_table.add_row("overtime_3", str(self.overtime_3))
        my_table.add_row("overtime_4", str(self.overtime_4))
        my_table.add_row("overtime_5", str(self.overtime_5))
        yield my_table


@dataclass
class HourTotals:
    last_week: timedelta
    this_week: timedelta
    total: timedelta

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield "[b]HourTotals:[/b]"
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
            "cGroupValue": self.value,
            "cGroupValueDescription": self.group_value_description,
            "cGroupCode": self.group_code,
            "iGroupUserType": self.group_user_type,
            "cGroupColor": self.group_color,
            "cDescription": self.description,
            "lClosed": self.closed,
        }
        d.update(self.gps.to_dict())
        return d

    def write_dict(self) -> Dict[str, Any]:
        d = {
            "iGroupNumber": self.group_number,
            "iGroupValueSeq": int(self.group_seq),
            "cGroupValue": self.value,
            "cGroupValueDescription": self.group_value_description,
            "isValid": True,
        }
        return d

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
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

    def __init__(self, value: int, name: str) -> None:
        self.value = value
        self.name = name
        self.logger = logging.getLogger('EntryCategoryGroup')

    def __getitem__(self, value: str) -> EntryCategory:
        with shelve.open(os.path.join('pay', self.name)) as options:
            if isinstance(value, int):
                value = str(value)
            try:
                return options[value]
            except KeyError:
                print(self)
                raise

    def __setitem__(self, value: str, option: EntryCategory) -> None:
        with shelve.open(os.path.join('pay', self.name)) as options:
            options[value] = option

    def __len__(self) -> int:
        l = 0
        with shelve.open(os.path.join('pay', self.name)) as options:
            l = len(options)
        return l

    def __contains__(self, value: str) -> bool:
        val = False
        with shelve.open(os.path.join('pay', self.name)) as options:
            val = value in options
        return val

    def update(self, options: List[EntryCategory]):
        with shelve.open(os.path.join('pay', self.name), writeback=True) as group_options:
            for option in options:
                if option.value in group_options:
                    old = group_options[option.value]
                    new = option
                    self.logger.warning(f'Overwriting {option.value}:'
                                        f'{old} ({old.value})'
                                        f'->{new} ({new.value})')
                group_options[option.value] = option

    def has_key(self, value: str) -> bool:
        return self.__contains__(value)

    def keys(self) -> List[str]:
        keys = []
        with shelve.open(os.path.join('pay', self.name)) as options:
            keys = [k for k in options.keys()]
        return keys

    def values(self) -> List[EntryCategory]:
        values = []
        with shelve.open(os.path.join('pay', self.name)) as options:
            values = [v for v in options.values()]
        return values

    def items(self) -> List[Tuple[str, EntryCategory]]:
        items = []
        with shelve.open(os.path.join('pay', self.name)) as options:
            items = [(k, v) for k, v in options.items()]
        return items

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
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
        with shelve.open(os.path.join('pay', self.name)) as options:
            for _, option in options.items():
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
class Schedule:
    """Representation of the schedule fields in a NOVATime timesheet."""
    schedule: str  # cSchedule
    is_schedule_premium: int  # isSchedulePremium
    is_schedule_premium_user_override: int  # isSchedulePremiumUserOverride
    schedule_hours: timedelta  # nScheduleHours
    grouping_string: str  # SchGroupingString

    def __init__(self,
                 schedule: str,  # cSchedule
                 is_schedule_premium: int,  # isSchedulePremium
                 is_schedule_premium_user_override: int,  # isSchedulePremiumUserOverride
                 schedule_hours: float,  # nScheduleHours
                 grouping_string: str  # SchGroupingString
                 ) -> None:
        self.schedule = schedule
        self.is_schedule_premium = is_schedule_premium
        self.is_schedule_premium_user_override = is_schedule_premium_user_override
        self.schedule_hours = timedelta(hours=schedule_hours)
        self.grouping_string = grouping_string

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
            'cSchedule': self.schedule,
            'isSchedulePremium': self.is_schedule_premium,
            'isSchedulePremiumUserOverride': self.is_schedule_premium_user_override,
            'nScheduleHours': self.schedule_hours.total_seconds()/3600.0,
            'SchGroupingString': self.grouping_string,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]Schedule:[/b] #{self.schedule}"
        my_table = Table("Attribute", "Value")
        my_table.add_row("is_schedule_premium", str(self.is_schedule_premium))
        my_table.add_row("is_schedule_premium_user_override",
                         str(self.is_schedule_premium_user_override))
        my_table.add_row("schedule_hours", str(self.schedule_hours))
        my_table.add_row("grouping_string", self.grouping_string)
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
    code_description: str  # cPayCodeDescription
    show_schedule: bool  # lShowSchedulePayCode
    non_compute_calc: bool  # lNonComputeCalcPayCode
    overtime_1: float  # nOT1Pay
    overtime_2: float  # nOT2Pay
    overtime_3: float  # nOT3Pay
    overtime_4: float  # nOT4Pay
    overtime_5: float  # nOT5Pay
    amount: float  # nPayAmount
    premium: float  # nPremiumPay
    regular: float  # nRegPay
    total: float  # nTotalPay

    def __str__(self) -> str:
        return f'{self.code}[{self.description}]'

    def to_dict(self) -> Dict[str, Any]:
        return {
            'nPayCode': self.code,
            'cExpCode': self.description,
            'nCodeType': self.code_type,
            'lPayCodeReadOnly': self.read_only,
            'cPayType': self.pay_type,
            'nPayPolicy': self.policy,
            'cPayPolicyDescription': self.policy_description,
            'nPayRate': self.rate,
            'cPayCodeDescription': self.code_description,
            'lShowSchedulePayCode': self.show_schedule,
            'lNonComputeCalcPayCode': self.non_compute_calc,
            'nOT1Pay': self.overtime_1,
            'nOT2Pay': self.overtime_2,
            'nOT3Pay': self.overtime_3,
            'nOT4Pay': self.overtime_4,
            'nOT5Pay': self.overtime_5,
            'nPayAmount': self.amount,
            'nPremiumPay': self.premium,
            'nRegPay': self.regular,
            'nTotalPay': self.total,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]PayCode:[/b] #{self.code}"
        my_table = Table("Attribute", "Value")
        my_table.add_row("description", self.description)
        my_table.add_row("code_type", str(self.code_type))
        my_table.add_row("read_only", str(self.read_only))
        my_table.add_row("pay_type", self.pay_type)
        my_table.add_row("policy", str(self.policy))
        my_table.add_row("policy_description", self.policy_description)
        my_table.add_row("rate", str(self.rate))
        my_table.add_row("code_description", self.code_description)
        my_table.add_row("show_schedule", str(self.show_schedule))
        my_table.add_row("non_compute_calc", str(self.non_compute_calc))
        my_table.add_row("overtime_1", str(self.overtime_1))
        my_table.add_row("overtime_2", str(self.overtime_2))
        my_table.add_row("overtime_3", str(self.overtime_3))
        my_table.add_row("overtime_4", str(self.overtime_4))
        my_table.add_row("overtime_5", str(self.overtime_5))
        my_table.add_row("amount", str(self.amount))
        my_table.add_row("premium", str(self.premium))
        my_table.add_row("regular", str(self.regular))
        my_table.add_row("total", str(self.total))
        yield my_table


@dataclass
class TimesheetEntryException:
    """Representation of the exception fields in a NOVATime timesheet."""
    # pylint: disable=too-many-instance-attributes
    # There is no good way to package these into fewer attributes, so pylint can blow it
    break_desc: str  # cBreakException
    meal_desc: str  # cMealException
    auto_deduct_meal: bool  # lAutoDeductMealException
    auto_meal_waived: bool  # lAutoMealWaivedException
    early_in: bool  # lEarlyInException
    early_out: bool  # lEarlyOutException
    early_out_grace: bool  # lEarlyOutGraceException
    late_out: bool  # lLateOutException
    late_out_to_meal: bool  # lLateOutToMealException
    meal_break_premium: bool  # lMealBreakPremiumException
    missing_punch: bool  # lMissingPunchException
    over_pay: bool  # lOverPayException
    overtime: bool  # lOvertimeException
    tardy: bool  # lTardyException
    tardy_grace: bool  # lTardyGraceException
    unauthorized_ot: bool  # lUnauthorizedOT
    unconfirmed_in: bool  # lUnconfirmedInException
    unconfirmed_in_punch: bool  # lUnconfirmedInPunch
    unconfirmed_out: bool  # lUnconfirmedOutException
    unconfirmed_out_punch: bool  # lUnconfirmedOutPunch
    unconfirmed_punch: bool  # lUnconfirmedPunchException
    under_pay: bool  # lUnderPayException
    unpaid_break: bool  # lUnpaidBreakException
    auto_meal_minutes: timedelta  # nAutoMealMinutes
    early_out_minutes: timedelta  # nEarlyOutMinutes
    long_meal: float  # nLongMeal
    meal_val_minutes: timedelta  # nMealValMinutes
    quantity_bad: float  # nQuantityBad
    quantity_good: float  # nQuantityGood
    tardy_minutes: timedelta  # nTardyMinutes

    def __init__(self,
                 break_desc: str,  # cBreakException
                 meal_desc: str,  # cMealException
                 auto_deduct_meal: bool,  # lAutoDeductMealException
                 auto_meal_waived: bool,  # lAutoMealWaivedException
                 early_in: bool,  # lEarlyInException
                 early_out: bool,  # lEarlyOutException
                 early_out_grace: bool,  # lEarlyOutGraceException
                 late_out: bool,  # lLateOutException
                 late_out_to_meal: bool,  # lLateOutToMealException
                 meal_break_premium: bool,  # lMealBreakPremiumException
                 missing_punch: bool,  # lMissingPunchException
                 over_pay: bool,  # lOverPayException
                 overtime: bool,  # lOvertimeException
                 tardy: bool,  # lTardyException
                 tardy_grace: bool,  # lTardyGraceException
                 unauthorized_ot: bool,  # lUnauthorizedOT
                 unconfirmed_in: bool,  # lUnconfirmedInException
                 unconfirmed_in_punch: bool,  # lUnconfirmedInPunch
                 unconfirmed_out: bool,  # lUnconfirmedOutException
                 unconfirmed_out_punch: bool,  # lUnconfirmedOutPunch
                 unconfirmed_punch: bool,  # lUnconfirmedPunchException
                 under_pay: bool,  # lUnderPayException
                 unpaid_break: bool,  # lUnpaidBreakException
                 auto_meal_minutes: float,  # nAutoMealMinutes
                 early_out_minutes: float,  # nEarlyOutMinutes
                 long_meal: float,  # nLongMeal
                 meal_val_minutes: float,  # nMealValMinutes
                 quantity_bad: float,  # nQuantityBad
                 quantity_good: float,  # nQuantityGood
                 tardy_minutes: float,  # nTardyMinutes
                 ) -> None:
        self.break_desc = break_desc
        self.meal_desc = meal_desc
        self.auto_deduct_meal = auto_deduct_meal
        self.auto_meal_waived = auto_meal_waived
        self.early_in = early_in
        self.early_out = early_out
        self.early_out_grace = early_out_grace
        self.late_out = late_out
        self.late_out_to_meal = late_out_to_meal
        self.meal_break_premium = meal_break_premium
        self.missing_punch = missing_punch
        self.over_pay = over_pay
        self.overtime = overtime
        self.tardy = tardy
        self.tardy_grace = tardy_grace
        self.unauthorized_ot = unauthorized_ot
        self.unconfirmed_in = unconfirmed_in
        self.unconfirmed_in_punch = unconfirmed_in_punch
        self.unconfirmed_out = unconfirmed_out
        self.unconfirmed_out_punch = unconfirmed_out_punch
        self.unconfirmed_punch = unconfirmed_punch
        self.under_pay = under_pay
        self.unpaid_break = unpaid_break
        self.auto_meal_minutes = timedelta(minutes=auto_meal_minutes)
        self.early_out_minutes = timedelta(minutes=early_out_minutes)
        self.long_meal = long_meal
        self.meal_val_minutes = timedelta(minutes=meal_val_minutes)
        self.quantity_bad = quantity_bad
        self.quantity_good = quantity_good
        self.tardy_minutes = timedelta(minutes=tardy_minutes)

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
            'cBreakException': self.break_desc,
            'cMealException': self.meal_desc,
            'lAutoDeductMealException': self.auto_deduct_meal,
            'lAutoMealWaivedException': self.auto_meal_waived,
            'lEarlyInException': self.early_in,
            'lEarlyOutException': self.early_out,
            'lEarlyOutGraceException': self.early_out_grace,
            'lLateOutException': self.late_out,
            'lLateOutToMealException': self.late_out_to_meal,
            'lMealBreakPremiumException': self.meal_break_premium,
            'lMissingPunchException': self.missing_punch,
            'lOverPayException': self.over_pay,
            'lOvertimeException': self.overtime,
            'lTardyException': self.tardy,
            'lTardyGraceException': self.tardy_grace,
            'lUnauthorizedOT': self.unauthorized_ot,
            'lUnconfirmedInException': self.unconfirmed_in,
            'lUnconfirmedInPunch': self.unconfirmed_in_punch,
            'lUnconfirmedOutException': self.unconfirmed_out,
            'lUnconfirmedOutPunch': self.unconfirmed_out_punch,
            'lUnconfirmedPunchException': self.unconfirmed_punch,
            'lUnderPayException': self.under_pay,
            'lUnpaidBreakException': self.unpaid_break,
            'nAutoMealMinutes': self.auto_meal_minutes.total_seconds()/60.0,
            'nEarlyOutMinutes': self.early_out_minutes.total_seconds()/60.0,
            'nLongMeal': self.long_meal,
            'nMealValMinutes': self.meal_val_minutes.total_seconds()/60.0,
            'nQuantityBad': self.quantity_bad,
            'nQuantityGood': self.quantity_good,
            'nTardyMinutes': self.tardy_minutes.total_seconds()/60.0,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield "[b]TimesheetEntryException:[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("overtime", str(self.overtime))
        my_table.add_row("unauthorized_ot", str(self.unauthorized_ot))
        my_table.add_row("tardy", str(self.tardy))
        my_table.add_row("tardy_grace", str(self.tardy_grace))
        my_table.add_row("tardy_minutes", str(self.tardy_minutes))
        my_table.add_row("early_in", str(self.early_in))
        my_table.add_row("early_out", str(self.early_out))
        my_table.add_row("early_out_grace", str(self.early_out_grace))
        my_table.add_row("early_out_minutes", str(self.early_out_minutes))
        my_table.add_row("late_out", str(self.late_out))
        my_table.add_row("under_pay", str(self.under_pay))
        my_table.add_row("over_pay", str(self.over_pay))
        my_table.add_row("unpaid_break", str(self.unpaid_break))
        my_table.add_row("break_desc", self.break_desc)
        my_table.add_row("meal_break_premium", str(self.meal_break_premium))
        my_table.add_row("meal_desc", self.meal_desc)
        my_table.add_row("long_meal", str(self.long_meal))
        my_table.add_row("late_out_to_meal", str(self.late_out_to_meal))
        my_table.add_row("auto_deduct_meal", str(self.auto_deduct_meal))
        my_table.add_row("auto_meal_waived", str(self.auto_meal_waived))
        my_table.add_row("auto_meal_minutes", str(self.auto_meal_minutes))
        my_table.add_row("meal_val_minutes", str(self.meal_val_minutes))
        my_table.add_row("missing_punch", str(self.missing_punch))
        my_table.add_row("unconfirmed_punch", str(self.unconfirmed_punch))
        my_table.add_row("unconfirmed_in_punch", str(
            self.unconfirmed_in_punch))
        my_table.add_row("unconfirmed_out_punch", str(
            self.unconfirmed_out_punch))
        my_table.add_row("unconfirmed_in", str(self.unconfirmed_in))
        my_table.add_row("unconfirmed_out", str(self.unconfirmed_out))
        my_table.add_row("quantity_bad", str(self.quantity_bad))
        my_table.add_row("quantity_good", str(self.quantity_good))
        yield my_table


@dataclass
class TimesheetEntryNote:
    """Representation of the note fields in a NOVATime timesheet."""
    in_out_expression: str  # cInOutExpression
    more_info: str  # cMoreInfo
    notes: str  # cNotes
    author: str  # cAuthor
    reason_code: str  # cReasonCode
    reason_color: str  # cReasonColor
    id: int  # iNoteSeq

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
            'cInOutExpression': self.in_out_expression,
            'cMoreInfo': self.more_info,
            'cNotes': self.notes,
            'cAuthor': self.author,
            'cReasonCode': self.reason_code,
            'cReasonColor': self.reason_color,
            'iNoteSeq': self.id,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield "[b]TimesheetEntryNote:[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("in_out_expression", str(self.in_out_expression))
        my_table.add_row("more_info", str(self.more_info))
        my_table.add_row("notes", str(self.notes))
        my_table.add_row("reason_code", str(self.reason_code))
        my_table.add_row("reason_color", str(self.reason_color))
        my_table.add_row("id", str(self.id))
        yield my_table


@dataclass
class TimesheetEntryStatus:
    """Representation of the status fields in a NOVATime timesheet."""
    approval: bool  # lApprovalStatus
    audit: bool  # lAudit
    calc_override: bool  # lCalcOverride
    carryover_expansion_or_changed: bool  # lCarryoverExpansionORChanged
    compute_non_calc: bool  # lComputeNonCalc
    elapsed_time: bool  # lElapsedTime
    has_last_change_day: bool  # lHasLstChgDay
    is_tga_record: bool  # lIsTGARecord
    pending: bool  # lPending
    pending_calc: bool  # lPendingCalc
    read_only: bool  # lReadOnly
    read_only_reason: str  # ReadOnlyReasonCodeDescription
    ref_time: bool  # lRefTime
    reversed: bool  # lReversed
    within_pay_period: bool  # lWithinPP
    auto_pay_no_delete: bool  # lAutoPayNoDelete
    approval_status: float  # nApprovalStatus
    calculate: float  # nCalculate
    reverse_status: float  # nReverseStatus

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
            'lApprovalStatus': self.approval,
            'lAudit': self.audit,
            'lCalcOverride': self.calc_override,
            'lCarryoverExpansionORChanged': self.carryover_expansion_or_changed,
            'lComputeNonCalc': self.compute_non_calc,
            'lElapsedTime': self.elapsed_time,
            'lHasLstChgDay': self.has_last_change_day,
            'lIsTGARecord': self.is_tga_record,
            'lPending': self.pending,
            'lPendingCalc': self.pending_calc,
            'lReadOnly': self.read_only,
            'ReadOnlyReasonCodeDescription': self.read_only_reason,
            'lRefTime': self.ref_time,
            'lReversed': self.reversed,
            'lWithinPP': self.within_pay_period,
            'lAutoPayNoDelete': self.auto_pay_no_delete,
            'nApprovalStatus': self.approval_status,
            'nCalculate': self.calculate,
            'nReverseStatus': self.reverse_status,
        }

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield "[b]TimesheetEntryStatus:[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("approval", str(self.approval))
        my_table.add_row("audit", str(self.audit))
        my_table.add_row("calc_override", str(self.calc_override))
        my_table.add_row("carryover_expansion_or_changed",
                         str(self.carryover_expansion_or_changed))
        my_table.add_row("compute_non_calc", str(self.compute_non_calc))
        my_table.add_row("elapsed_time", str(self.elapsed_time))
        my_table.add_row("has_last_change_day", str(self.has_last_change_day))
        my_table.add_row("is_tga_record", str(self.is_tga_record))
        my_table.add_row("pending", str(self.pending))
        my_table.add_row("pending_calc", str(self.pending_calc))
        my_table.add_row("read_only", str(self.read_only))
        my_table.add_row("ref_time", str(self.ref_time))
        my_table.add_row("reversed", str(self.reversed))
        my_table.add_row("within_pay_period", str(self.within_pay_period))
        my_table.add_row("auto_pay_no_delete", str(self.auto_pay_no_delete))
        my_table.add_row("approval_status", str(self.approval_status))
        my_table.add_row("calculate", str(self.calculate))
        my_table.add_row("reverse_status", str(self.reverse_status))
        my_table.add_row("read_only_reason", str(self.read_only_reason))
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

    def __init__(self,
                 punch: Union[str, arrow.arrow.Arrow],
                 adjust: Union[str, arrow.arrow.Arrow],
                 modified: bool,
                 gps: GPS,
                 site: str,
                 og: Union[str, arrow.arrow.Arrow],
                 net_chk_fail: bool,
                 expression: str,
                 expression_save: str,
                 timezone: Union[str, arrow.arrow.Arrow],
                 recording: Any,
                 ) -> None:
        self.modified = modified
        self.gps = gps
        self.site = site
        self.net_chk_fail = net_chk_fail
        self.expression = expression
        self.expression_save = expression_save
        self.recording = recording

        self.punch = parse_punch(punch)
        if isinstance(adjust, str):
            adjust = arrow.get(adjust, "HH:mmA")
            self.adjust = arrow.get(datetime.combine(
                self.punch.date(), adjust.time()))

        self.og = parse_punch(og)
        self.timezone = parse_punch(timezone)

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]Punch: {self.punch}[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("adjust", str(self.adjust))
        my_table.add_row("modified", str(self.modified))
        my_table.add_row("gps", str(self.gps))
        my_table.add_row("site", self.site)
        my_table.add_row("og", str(self.og))
        my_table.add_row("net_chk_fail", str(self.net_chk_fail))
        my_table.add_row("expression", self.expression)
        my_table.add_row("expression_save", self.expression_save)
        my_table.add_row("timezone", str(self.timezone))
        my_table.add_row("recording", str(self.recording))
        yield my_table


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

    pay_code: PayCode
    # nOT1Pay, nOT2Pay, nOT3Pay, nOT4Pay, nOT5Pay, nPayAmount, nPremiumPay, nRegPay, nTotalPay,
    # cExpCode, cPayType, cPayPolicyDescription, cPayCodeDescription, lShowSchedulePayCode,
    # lNonComputeCalcPayCode, lPayCodeReadOnly, nPayCode, nCodeType, nPayPolicy, nPayRate

    categories: List[EntryCategory]     # GroupValueList
    accessible_group_list: List[EntryCategory]  # AccessibleGroupList
    invalid_group_list: List[EntryCategory]  # InvalidGroupList
    grouping: EntryCategory             # Grouping
    exceptions: TimesheetEntryException
    shift_expression: str

    adjustment_date: arrow.arrow.Arrow  # dAdjustmentDate
    assign_id: str  # cAssignID
    comp_hours: timedelta  # nCompHours
    daily_hours: timedelta  # nDailyHours
    daily_total_hours: timedelta  # nDailyTotalHours
    date_key: arrow.arrow.Arrow  # DateKey
    group_code: str  # cGroupCode
    group_level: str  # cGroupLevel
    grouping_string: str  # GroupingString
    last_modified: arrow.arrow.Arrow  # tLastModified
    reverse_date: arrow.arrow.Arrow  # dReverseDate
    shift_expression: str  # cShiftExpression
    total_hours: timedelta  # nTotalHours
    week_group_string: str  # WeekGroupString
    weekly_hours: timedelta  # nWeeklyHours
    weekly_hours_total: timedelta  # nWeeklyTotalHours
    work_hours: timedelta  # nWorkHours

    expected_meal_times: Dict[int, DatePeriod]  # ExpectedMealTimes
    overtime_hours: HoursGroup  # nOT*Hours
    redirect_overtime_hours: HoursGroup  # nRedirectOT*Hours
    comp_overtime_hours: HoursGroup  # nCompOT*Hours
    raw_comp_overtime_hours: HoursGroup  # nRawCompOT*Hours

    # cSchedule, isSchedulePremium, isSchedulePremiumUserOverride, nScheduleHours, SchGroupingString
    schedule: Schedule

    # cInOutExpression, cMoreInfo, cNotes, cAuthor, cReasonCode, cReasonColor, iNoteSeq
    note: TimesheetEntryNote

    status: TimesheetEntryStatus
    # lApprovalStatus, lAudit, lCalcOverride, lCarryoverExpansionORChanged, lComputeNonCalc,
    # lElapsedTime, lHasLstChgDay, lIsTGARecord, lPending, lPendingCalc, lReadOnly, lRefTime,
    # lReversed, lWithinPP, lAutoPayNoDelete, nApprovalStatus, nCalculate, nReverseStatus

    carryover_expansion_override: Any  # CarryoverExpansionOverride
    overtime_total_hours_one_punch: Any  # OTTotalHoursOnePunch
    record_type: Any  # RecordType

    def __init__(self, entry: Dict[str, Any]) -> None:
        self.sheet_sequence = entry["iTimesheetSeq"]

        self.week_group_string = entry["WeekGroupString"]

        self.pay_period = DatePeriod(start=entry["dPayPeriodStart"],
                                     end=entry["dPayPeriodEnd"])

        self.work_period = DatePeriod(start=entry["dWorkPeriodStartDate"],
                                      end=entry["dWorkPeriodEndDate"])
        try:
            self.entry_sequence = entry["iTimeSeq"]
        except KeyError:
            self.entry_sequence = -1

        self.employee = User(employee_seq=entry["iEmployeeSeq"],
                             username=entry["cEmployeeID"],
                             first_name=entry["cEmployeeFirstName"],
                             last_name=entry["cEmployeeLastName"],
                             full_name=entry["cEmployeeFullName"],
                             password=None)

        self.punch_date = parse_punch(entry["dPunchDate"])
        self.work_date = parse_punch(entry["dWorkDate"])
        self.punch_date_time = parse_punch(entry["tPunchDateTime"])

        self.punch_in = Punch(
            punch=entry["dIn"],
            adjust=entry["nAdjustIn"],
            modified=entry["lInMod"],
            gps=entry["cInGPS"],
            site=entry["cSiteIn"],
            og=entry["dOGIn"],
            net_chk_fail=entry["lInNetChkFail"],
            expression=entry["cInExpression"],
            expression_save=entry["cInExpressionSave"],
            timezone=entry["nTZIn"],
            recording=entry["mInRecording"]
        )
        if entry['dOut'] is None:
            self.punch_out = None
        else:
            self.punch_out = Punch(
                punch=entry["dOut"],
                adjust=entry["nAdjustOut"],
                modified=entry["lOutMod"],
                gps=entry["cOutGPS"],
                site=entry["cSiteOut"],
                og=entry["dOGOut"],
                net_chk_fail=entry["lOutNetChkFail"],
                expression=entry["cOutExpression"],
                expression_save=entry["cOutExpressionSave"],
                timezone=entry["nTZOut"],
                recording=entry["mOutRecording"]
            )
        self.date_key = entry["DateKey"]
        self.work_hours = timedelta(hours=entry["nWorkHours"])
        self.total_hours = timedelta(hours=entry["nTotalHours"])
        self.comp_hours = timedelta(hours=entry["nCompHours"])
        self.daily_hours = timedelta(hours=entry["nDailyHours"])
        self.daily_total_hours = timedelta(hours=entry["nDailyTotalHours"])
        self.weekly_hours = timedelta(hours=entry["nWeeklyHours"])
        self.weekly_hours_total = timedelta(hours=entry["nWeeklyTotalHours"])

        self.overtime_hours = HoursGroup(ot_type='',
                                         overtime_1=entry["nOT1Hours"],
                                         overtime_2=entry["nOT2Hours"],
                                         overtime_3=entry["nOT3Hours"],
                                         overtime_4=entry["nOT4Hours"],
                                         overtime_5=entry["nOT5Hours"],
                                         )

        self.comp_overtime_hours = HoursGroup(ot_type='Comp',
                                              overtime_1=entry["nCompOT1Hours"],
                                              overtime_2=entry["nCompOT2Hours"],
                                              overtime_3=entry["nCompOT3Hours"],
                                              overtime_4=entry["nCompOT4Hours"],
                                              overtime_5=entry["nCompOT5Hours"],
                                              )

        self.raw_comp_overtime_hours = HoursGroup(ot_type='RawComp',
                                                  overtime_1=entry["nRawCompOT1Hours"],
                                                  overtime_2=entry["nRawCompOT2Hours"],
                                                  overtime_3=entry["nRawCompOT3Hours"],
                                                  overtime_4=entry["nRawCompOT4Hours"],
                                                  overtime_5=entry["nRawCompOT5Hours"],
                                                  )

        self.redirect_overtime_hours = HoursGroup(ot_type='Redirect',
                                                  overtime_1=entry["nRedirectOT1Hours"],
                                                  overtime_2=entry["nRedirectOT2Hours"],
                                                  overtime_3=entry["nRedirectOT3Hours"],
                                                  overtime_4=entry["nRedirectOT4Hours"],
                                                  overtime_5=entry["nRedirectOT5Hours"],
                                                  )

        self.pay_code = PayCode(
            code=entry["nPayCode"],
            description=entry["cExpCode"],
            code_type=entry["nCodeType"],
            read_only=entry["lPayCodeReadOnly"],
            pay_type=entry["cPayType"],
            policy=entry["nPayPolicy"],
            policy_description=entry["cPayPolicyDescription"],
            rate=entry["nPayRate"],
            code_description=entry["cPayCodeDescription"],
            show_schedule=entry["lShowSchedulePayCode"],
            non_compute_calc=entry["lNonComputeCalcPayCode"],
            overtime_1=entry["nOT1Pay"],
            overtime_2=entry["nOT2Pay"],
            overtime_3=entry["nOT3Pay"],
            overtime_4=entry["nOT4Pay"],
            overtime_5=entry["nOT5Pay"],
            amount=entry["nPayAmount"],
            premium=entry["nPremiumPay"],
            regular=entry["nRegPay"],
            total=entry["nTotalPay"],
        )

        self.categories = []
        if entry["GroupValueList"]:
            for group in entry["GroupValueList"]:
                self.categories.append(EntryCategory(group))

        self.accessible_group_list = []
        if entry["AccessibleGroupList"]:
            for group in entry["AccessibleGroupList"]:
                self.accessible_group_list.append(EntryCategory(group))

        self.invalid_group_list = []
        if entry["InvalidGroupList"]:
            for group in entry["InvalidGroupList"]:
                self.invalid_group_list.append(EntryCategory(group))

        self.grouping = EntryCategory(entry["Grouping"])
        self.grouping_string = entry["GroupingString"]
        self.group_level = entry["cGroupLevel"]
        self.group_code = entry["cGroupCode"]

        self.expected_meal_times = {}
        if entry["ExpectedMealTimes"]:
            for meal_time in entry["ExpectedMealTimes"]:
                self.expected_meal_times[meal_time["iIndex"]] = DatePeriod(
                    start=meal_time["tStartTime"],
                    end=meal_time["tEndTime"]
                )

        self.exceptions = TimesheetEntryException(
            break_desc=entry['cBreakException'],
            meal_desc=entry['cMealException'],
            auto_deduct_meal=entry['lAutoDeductMealException'],
            auto_meal_waived=entry['lAutoMealWaivedException'],
            early_in=entry['lEarlyInException'],
            early_out=entry['lEarlyOutException'],
            early_out_grace=entry['lEarlyOutGraceException'],
            late_out=entry['lLateOutException'],
            late_out_to_meal=entry['lLateOutToMealException'],
            meal_break_premium=entry['lMealBreakPremiumException'],
            missing_punch=entry['lMissingPunchException'],
            over_pay=entry['lOverPayException'],
            overtime=entry['lOvertimeException'],
            tardy=entry['lTardyException'],
            tardy_grace=entry['lTardyGraceException'],
            unauthorized_ot=entry['lUnauthorizedOT'],
            unconfirmed_in=entry['lUnconfirmedInException'],
            unconfirmed_in_punch=entry['lUnconfirmedInPunch'],
            unconfirmed_out=entry['lUnconfirmedOutException'],
            unconfirmed_out_punch=entry['lUnconfirmedOutPunch'],
            unconfirmed_punch=entry['lUnconfirmedPunchException'],
            under_pay=entry['lUnderPayException'],
            unpaid_break=entry['lUnpaidBreakException'],
            auto_meal_minutes=entry['nAutoMealMinutes'],
            early_out_minutes=entry['nEarlyOutMinutes'],
            long_meal=entry['nLongMeal'],
            meal_val_minutes=entry['nMealValMinutes'],
            quantity_bad=entry['nQuantityBad'],
            quantity_good=entry['nQuantityGood'],
            tardy_minutes=entry['nTardyMinutes'],
        )

        self.status = TimesheetEntryStatus(
            approval=entry["lApprovalStatus"],
            audit=entry["lAudit"],
            calc_override=entry["lCalcOverride"],
            carryover_expansion_or_changed=entry["lCarryoverExpansionORChanged"],
            compute_non_calc=entry["lComputeNonCalc"],
            elapsed_time=entry["lElapsedTime"],
            has_last_change_day=entry["lHasLstChgDay"],
            is_tga_record=entry["lIsTGARecord"],
            pending=entry["lPending"],
            pending_calc=entry["lPendingCalc"],
            read_only=entry["lReadOnly"],
            ref_time=entry["lRefTime"],
            reversed=entry["lReversed"],
            within_pay_period=entry["lWithinPP"],
            auto_pay_no_delete=entry["lAutoPayNoDelete"],
            approval_status=entry["nApprovalStatus"],
            calculate=entry["nCalculate"],
            reverse_status=entry["nReverseStatus"],
            read_only_reason=entry["ReadOnlyReasonCodeDescription"]
        )

        self.schedule = Schedule(
            schedule=entry["cSchedule"],
            is_schedule_premium=entry["isSchedulePremium"],
            is_schedule_premium_user_override=entry["isSchedulePremiumUserOverride"],
            schedule_hours=entry["nScheduleHours"],
            grouping_string=entry["SchGroupingString"]
        )

        self.note = TimesheetEntryNote(
            in_out_expression=entry["cInOutExpression"],
            more_info=entry["cMoreInfo"],
            notes=entry["cNotes"],
            reason_code=entry["cReasonCode"],
            reason_color=entry["cReasonColor"],
            id=entry["iNoteSeq"],
            author=entry["cAuthor"],
        )

        self.overtime_total_hours_one_punch = entry["OTTotalHoursOnePunch"]
        self.shift_expression = entry["cShiftExpression"]
        self.record_type = entry["RecordType"]
        self.adjustment_date = entry["dAdjustmentDate"]
        self.reverse_date = entry["dReverseDate"]
        self.last_modified = entry["tLastModified"]
        self.carryover_expansion_override = entry["CarryoverExpansionOverride"]
        self.assign_id = entry["cAssignID"]
        self.copy_color = 'darkgrey'

    def to_dict(self) -> Dict[str, Any]:
        d = {
            "iTimesheetSeq": self.sheet_sequence,
            "WeekGroupString": self.week_group_string,
            "iTimeSeq": self.entry_sequence,
            "dPunchDate": self.punch_date,
            "dWorkDate": self.work_date,
            "tPunchDateTime": self.punch_date_time,
            "DateKey": self.date_key,
            "nWorkHours": self.work_hours.total_seconds()/3600.0,
            "nTotalHours": self.total_hours.total_seconds()/3600.0,
            "nCompHours": self.comp_hours.total_seconds()/3600.0,
            "nDailyHours": self.daily_hours.total_seconds()/3600.0,
            "nDailyTotalHours": self.daily_total_hours.total_seconds()/3600.0,
            "nWeeklyHours": self.weekly_hours.total_seconds()/3600.0,
            "nWeeklyTotalHours": self.weekly_hours_total.total_seconds()/3600.0,
            "GroupingString": self.grouping_string,
            "cGroupLevel": self.group_level,
            "cGroupCode": self.group_code,
            "OTTotalHoursOnePunch": self.overtime_total_hours_one_punch,
            "cShiftExpression": self.shift_expression,
            "RecordType": self.record_type,
            "dAdjustmentDate": self.adjustment_date,
            "dReverseDate": self.reverse_date,
            "tLastModified": self.last_modified,
            "CarryoverExpansionOverride": self.carryover_expansion_override,
            "cAssignID": self.assign_id,
            "dPayPeriodStart": self.pay_period.start,
            "dPayPeriodEnd": self.pay_period.end,
            "dWorkPeriodStartDate": self.work_period.start,
            "dWorkPeriodEndDate": self.work_period.end,
            "iEmployeeSeq": self.employee.employee_seq,
            "cEmployeeID": self.employee.username,
            "cEmployeeFirstName": self.employee.first_name,
            "cEmployeeLastName": self.employee.last_name,
            "cEmployeeFullName": self.employee.full_name,
            "dIn": self.punch_in.punch,
            "nAdjustIn": self.punch_in.adjust,
            "lInMod": self.punch_in.modified,
            "cInGPS": self.punch_in.gps,
            "cSiteIn": self.punch_in.site,
            "dOGIn": self.punch_in.og,
            "lInNetChkFail": self.punch_in.net_chk_fail,
            "cInExpression": self.punch_in.expression,
            "cInExpressionSave": self.punch_in.expression_save,
            "nTZIn": self.punch_in.timezone,
            "mInRecording": self.punch_in.recording,
            "dOut": self.punch_out.punch,
            "nAdjustOut": self.punch_out.adjust,
            "lOutMod": self.punch_out.modified,
            "cOutGPS": self.punch_out.gps,
            "cSiteOut": self.punch_out.site,
            "dOGOut": self.punch_out.og,
            "lOutNetChkFail": self.punch_out.net_chk_fail,
            "cOutExpression": self.punch_out.expression,
            "cOutExpressionSave": self.punch_out.expression_save,
            "nTZOut": self.punch_out.timezone,
            "mOutRecording": self.punch_out.recording,
        }
        d.update(self.overtime_hours.to_dict())
        d.update(self.comp_overtime_hours.to_dict())
        d.update(self.raw_comp_overtime_hours.to_dict())
        d.update(self.redirect_overtime_hours.to_dict())
        d.update(self.pay_code.to_dict())
        d["GroupValueList"] = [category.to_dict()
                               for category in self.categories] or None
        d["AccessibleGroupList"] = [category.to_dict()
                                    for category in self.accessible_group_list] or None
        d["InvalidGroupList"] = [category.to_dict()
                                 for category in self.invalid_group_list] or None
        d["Grouping"] = self.grouping.to_dict()

        d["ExpectedMealTimes"] = [{"iIndex": idx,
                                   "tStartTime": period.start,
                                   "tEndTime": period.end} for idx, period in self.expected_meal_times.items()] or None

        d.update(self.exceptions.to_dict())
        d.update(self.status.to_dict())
        d.update(self.schedule.to_dict())
        d.update(self.note.to_dict())

    def write_dict(self) -> Dict[str, Any]:
        d = {
            "iTimeSeq": self.entry_sequence,
            "iEmployeeSeq": self.employee.employee_seq,
            "copyColor": self.copy_color,
            "lUnauthorizedOT": self.exceptions.unauthorized_ot,
            "lCalcOverride": self.status.calc_override,
            "iTimesheetSeq": self.sheet_sequence,
            "cIn_Mod": self.punch_in.modified or "",
            "cOut_Mod": self.punch_out.modified or "",
            "isUnEditable": False,
            # "$$hashKey": "object:264",
            "payCodeDisplayText": "",
            "dPunchDate": self.punch_date.to('utc').format(WRITE_ENTRY_FORMAT),
            "dWorkDate": self.work_date.to('utc').format(WRITE_ENTRY_FORMAT),
            "nPayCode": self.pay_code.code,
            "nCodeType": self.pay_code.code_type,
            "nCalculate": self.status.calculate,
            "lRefTime": self.status.ref_time,
            "lChanged": True,
            "dIn": self.punch_in.punch.to('utc').format(WRITE_ENTRY_FORMAT),
            "dOut": self.punch_out.punch.to('utc').format(WRITE_ENTRY_FORMAT),
        }
        d["GroupValueList"] = [category.write_dict()
                               for category in self.categories[:4]] or None
        d.update({
            "focusGroup": 0,
            "cNotes": self.note.notes,
            "lHasError": False,
            "lAllowOverlapHours": False,
            "isPAPaycode": False,
            "lOverlappingTime": False,
            "lPendingCalc": True
        })
        return d

    def __rich_console__(self, console: Console, options: ConsoleOptions) -> RenderResult:
        # pylint: disable=unused-argument
        yield f"[b]TimesheetEntry: {self.punch_date}[/b]"
        my_table = Table("Attribute", "Value")
        my_table.add_row("sheet_sequence", str(self.sheet_sequence))
        my_table.add_row("pay_period", str(self.pay_period))
        my_table.add_row("work_period", str(self.work_period))
        my_table.add_row("entry_sequence", str(self.entry_sequence))
        my_table.add_row("employee", str(self.employee))
        my_table.add_row("work_date", str(self.work_date))
        my_table.add_row("punch_date_time", str(self.punch_date_time))
        my_table.add_row("punch_in", str(self.punch_in))
        my_table.add_row("punch_out", str(self.punch_out))
        my_table.add_row("pay_code", str(self.pay_code))
        my_table.add_row("categories", str(self.categories))
        my_table.add_row("accessible_group_list",
                         str(self.accessible_group_list))
        my_table.add_row("invalid_group_list", str(self.invalid_group_list))
        my_table.add_row("grouping", str(self.grouping))
        my_table.add_row("exceptions", str(self.exceptions))
        my_table.add_row("shift_expression", str(self.shift_expression))
        my_table.add_row("adjustment_date", str(self.adjustment_date))
        my_table.add_row("assign_id", str(self.assign_id))
        my_table.add_row("comp_hours", str(self.comp_hours))
        my_table.add_row("daily_hours", str(self.daily_hours))
        my_table.add_row("daily_total_hours", str(self.daily_total_hours))
        my_table.add_row("date_key", str(self.date_key))
        my_table.add_row("group_code", str(self.group_code))
        my_table.add_row("group_level", str(self.group_level))
        my_table.add_row("grouping_string", str(self.grouping_string))
        my_table.add_row("last_modified", str(self.last_modified))
        my_table.add_row("reverse_date", str(self.reverse_date))
        my_table.add_row("shift_expression", str(self.shift_expression))
        my_table.add_row("total_hours", str(self.total_hours))
        my_table.add_row("week_group_string", str(self.week_group_string))
        my_table.add_row("weekly_hours", str(self.weekly_hours))
        my_table.add_row("weekly_hours_total", str(self.weekly_hours_total))
        my_table.add_row("work_hours", str(self.work_hours))
        my_table.add_row("expected_meal_times", str(self.expected_meal_times))
        my_table.add_row("overtime_hours", str(self.overtime_hours))
        my_table.add_row("redirect_overtime_hours",
                         str(self.redirect_overtime_hours))
        my_table.add_row("comp_overtime_hours", str(self.comp_overtime_hours))
        my_table.add_row("raw_comp_overtime_hours",
                         str(self.raw_comp_overtime_hours))
        my_table.add_row("schedule", str(self.schedule))
        my_table.add_row("note", str(self.note))
        my_table.add_row("status", str(self.status))
        my_table.add_row("carryover_expansion_override",
                         str(self.carryover_expansion_override))
        my_table.add_row("overtime_total_hours_one_punch",
                         str(self.overtime_total_hours_one_punch))
        my_table.add_row("record_type", str(self.record_type))
        yield my_table


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
        # pylint: disable=redefined-builtin
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
        loginrequest = self._session.get(
            uri, params={'CID': self.cid}, verify=False)
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
                                      data=login_data, verify=False)
        if not response.ok:
            raise ValueError(
                f'Bad response: {response.status_code} - {response.reason}')

        # the SessionVariable contains several pieces of information we need for future requests;
        #  add them to the session headers and the secrets
        self.logger.debug(
            f'Asking for NOVATime user details for {self.user.username}')
        user_data_request = self._session.get(
            url=f'https://{self.api_url}/SessionVariable', verify=False)

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
            url=f'https://{self.api_url}/employee/{self.user.employee_seq}', verify=False).json()

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

        response = self._session.get(uri, params=parameters, verify=False)
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
        response = self._session.get(uri, verify=False)
        for group in response.json()['GROUPLIST']:
            self.groups[group['cGroupCaption']] = EntryCategoryGroup(
                value=group['iGroupNumber'],
                name=group['cGroupCaption']
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
        items_per_page = 100
        page = 1
        group_options = self.groups[group]
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
                response = self._session.get(
                    uri, params=parameters, verify=False)
                group_data = response.json()
                if group_data['_errorCode'] != 1:
                    raise ValueError(
                        f"Error getting group {group_options.name}: {group_data['_errorCode']} - "
                        f"{group_data['_errorDescription']}")
                items = group_data['Data']['ItemTotal']
                if len(group_options) == items:
                    self.logger.info(f'Already have all {items} for {group}')
                    progress.update(group_task, completed=True)
                    break
                if page == 1:
                    progress.update(group_task, total=items)
                    progress.start_task(group_task)
                progress.update(group_task, advance=len(
                    group_data['Data']['PagedList']))
                options = list(
                    map(EntryCategory, group_data['Data']['PagedList']))
                group_options.update(options)
                if len(group_options) < items:
                    page += 1
                    parameters['CurrentPage'] = str(page)
                    continue
                self.logger.debug(
                    f'Got {len(group_options)} of {items} for {group_options.name}')
            self.groups[group_options.name] = group_options
        return group_options

    def write_entries(self, entries: List[TimesheetEntry]) -> None:
        # build uri, parameters, and headers
        uri = f'https://{self.api_url}/timesheetdetail'
        parameters = {
            'AccessSeq': self.user.access_seq,
            'EmployeeSeq': self.user.employee_seq,
            'UserSeq': self.user.user_seq,
            "lOverWritePayPeriod": False,
            "readOnlyTimerecSeqList": "",
            "PolicyGroup": "",
            "showDaily": False
        }
        headers = {
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json;charset=utf-8",
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-gpc": "1",
            "userseq": str(self.user.user_seq) or '',
        }
        new_entries = [entry.write_dict() for entry in entries]
        response = self._session.post(uri, params=parameters,
                                      json=new_entries, headers=headers, verify=False)
        if not response.ok:
            raise ValueError(
                f'Bad response: {response.status_code} - {response.reason}')
        if not response.json()['_errorCode'] == 1:
            raise ValueError(f'API Error: {response.json()}')

        for result, entry in zip(response.json()['DataList'], entries):
            if result['Success']:
                self.logger.debug(f'Wrote {entry} successfully')
            else:
                self.logger.warning(f'Error writing {entry}: {result}')


class Timesheet:
    """Represent a NOVATime timesheet for a pay period"""
    entries = Dict[arrow.arrow.Arrow, List[TimesheetEntry]]
    hours: HourTotals
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

    def make_timesheet_report(self):
        """Print a little timesheet report."""

        print(f'Pay period from {self.pay_period.start.date()}'
              f' to {self.pay_period.end.date()}:')
        print('\tExceptions:')
        for date, entries in self.entries.items():
            print(f'\t\t{date.format("dddd, MMMM DD, YYYY")}:')
            for entry in entries:
                print(f'\t\t\t{entry.exceptions}')
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

    def get_times(self, raw_timesheet):
        """Retrieve the daily hours during the pay period."""
        self.entries = {}
        self.hours = HourTotals(
            last_week=timedelta(hours=0),
            this_week=timedelta(hours=0),
            total=timedelta(hours=0)
        )
        for raw_entry in raw_timesheet:
            entry = TimesheetEntry(raw_entry)
            if entry.punch_date not in self.entries:
                self.entries[entry.punch_date] = []
            self.entries[entry.punch_date].append(entry)

            if self.is_this_week(entry.punch_date):
                self.hours.this_week += entry.daily_total_hours
            else:
                self.hours.last_week += entry.daily_total_hours
            self.hours.total += entry.daily_total_hours

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
        for date, entries in self.entries.items():
            if date == today:
                today = entries
                # found today's entry
                break

        if isinstance(today, arrow.arrow.Arrow):
            # no entry today
            return None, None
        for entry in today:
            clock_in = entry.punch_in.punch
            if entry.exceptions.missing_punch:
                # we are clocked in, figure out a good clock-out time
                clock_out = clock_in + remaining
                if remaining > timedelta(hours=8):
                    # long enough shift we need a lunch
                    clock_out = clock_out.shift(
                        seconds=entry.exceptions.auto_meal_minutes.total_seconds())
            else:
                # not clocked in
                clock_out = None
        return clock_in, clock_out


if __name__ == '__main__':
    n = NOVATime()
    n.login()
    n.get_timesheet()
    n.timesheet.make_timesheet_report()
    for date, entries in n.timesheet.entries.items():
        print(date)
        for entry in entries:
            print(entry)
