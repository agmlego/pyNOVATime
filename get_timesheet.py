# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL

import configparser
import json
import os
import sys
from datetime import timedelta
from os.path import isdir, isfile, join

import arrow
import requests
from bs4 import BeautifulSoup

PARAM_DATE_FORMAT = 'ddd MMM DD YYYY'
PUNCH_FORMAT = 'MM/DD/YYYY HH:mm:ss'
NOVA_DATE_FORMAT = 'MM/DD/YYYY'


def format_td(td):
    '''Format a `datetime.timedelta` as hours and minutes'''
    hours, remainder = divmod(td.total_seconds(), 3600)
    minutes, _ = divmod(remainder, 60)
    return f'{hours:.0f}:{minutes:02.0f}'


def get_current_pay_period(last_pay):
    '''Calculate the bounds of the pay period.

    Our pay periods are two weeks long, ending one week after a paycheck:
        Always begins on a Monday (weekday = 0)
        Always ends on a Sunday (weekday = 6)

    Keyword arguments:
        last_pay -- an `arrow` indicating the last paycheck

    Returns a tuple with:
        start_date -- an `arrow` indicating the beginning of the pay period
        end_date -- an `arrow` indicating the end of the pay period

    '''
    end_date = last_pay.shift(weeks=1, weekday=6)

    # pay period starts two weeks before the end, on a Monday (=0)
    start_date = end_date.shift(weeks=-2, weekday=0)
    return (start_date, end_date)


def get_timesheet(session, secrets, start_date, end_date):
    '''Log in and download the timesheet.

    Here there be dragons. This does no error-checking, and is brittle as most scrapers are.

    Keyword arguments:
        session -- a logged-in `requests.Session` from `login()`
        secrets -- a `ConfigParser` for the secrets file (see `secrets.ini.example`)
        start_date -- an `arrow` indicating the beginning of the pay period
        end_date -- an `arrow` indicating the end of the pay period

    Returns a `requests.models.Response` from the site, hopefully containing the timesheet as JSON

    '''
    # build uri, parameters, and headers
    cid = secrets['uri']['cid']
    host = secrets['uri']['host']
    page = secrets['uri']['page']

    uri = f"https://{host}/{page}/{cid}/timesheetdetail"
    parameters = {
        'AccessSeq': secrets['user']['AccessSeq'],
        'EmployeeSeq': secrets['user']['EmployeeSeq'],
        'UserSeq': secrets['user']['UserSeq'],
        'StartDate': start_date.format(PARAM_DATE_FORMAT),
        'EndDate': end_date.format(PARAM_DATE_FORMAT),
        'CustomDateRange': False,
        'ShowOneMoreDay': False,
        'EmployeeSeqList': '',
        'DailyDate': start_date.format(PARAM_DATE_FORMAT),
        'ForceAbsent': False,
        'PolicyGroup': ''
    }

    r = session.get(uri, params=parameters)

    return r


def parse_punch(punch_str):
    '''Parse a punch entry into an `arrow`'''
    return arrow.get(punch_str, PUNCH_FORMAT).replace(tzinfo='America/Detroit')


def parse_date(date_str):
    '''Parse a date entry into an `arrow`'''
    return arrow.get(date_str, NOVA_DATE_FORMAT).replace(tzinfo='America/Detroit')


def get_exceptions(timesheet):
    '''Retrieve any exceptions during the pay period.

    There are likely lots of them, and many are innocuous.

    Keyword arguments:
        timesheet -- the 'DataList' list from the webpage response JSON

    Returns a dict with an `arrow` key for each day with exceptions mapped to a dict of the exception name and value

    '''
    exceptions = {}
    for entry in timesheet:
        punch_date = parse_date(entry['dPunchDate'])
        exceptions[punch_date] = {key: entry[key]
                                  for key in entry if 'Exception' in key and entry[key]}
        if not exceptions[punch_date]:
            del exceptions[punch_date]
    return exceptions


def get_times(timesheet):
    '''Retrieve the daily hours during the pay period.

    Keyword arguments:
        timesheet -- the 'DataList' list from the webpage response JSON

    Returns a dict with an `arrow` key for each day and `timedelta` hours worked,
        as well as a 'last_week' and 'this_week' and 'total' totals for those periods

    '''
    hours = {}
    hours['last_week'] = timedelta(hours=0)
    hours['this_week'] = timedelta(hours=0)
    hours['total'] = timedelta(hours=0)
    for entry in timesheet:
        punch_date = parse_date(entry['dPunchDate'])
        hours[punch_date] = timedelta(hours=entry['nDailyHours'])
        if is_this_week(punch_date):
            hours['this_week'] += hours[punch_date]
        else:
            hours['last_week'] += hours[punch_date]
        hours['total'] += hours[punch_date]
    return hours


def is_this_week(date):
    '''Report whether a date is in this week'''
    this_sunday = arrow.now().shift(weekday=6).floor('day')
    last_sunday = this_sunday.shift(weeks=-1).floor('day')
    return last_sunday <= date <= this_sunday


def predict_clock_out(timesheet, remaining):
    '''Try to predict an appropriate clock-out time.

    Keyword arguments:
        timesheet -- the 'DataList' list from the webpage response JSON
        remaining -- a `timedelta` indicating the remaining hours this week

    Returns a tuple with:
        clock_in -- an `arrow` if the day has a clock-in time, or None if not
        clock_out -- an `arrow` of clock_in plus remaining plus lunch if remaining is longer than 8 hours;
            however, if there is no "missing punch" exception (AKA we are not clocked in!), this will be None

    '''
    today = arrow.now().floor('day')
    for entry in timesheet:
        if parse_date(entry['dPunchDate']) == today:
            today = entry
            # found today's entry
            break

    if type(today) is arrow.arrow.Arrow:
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


def _ado_to_dict(datalist):
    """
    Clean up the ADO DataList objects that sometimes come from NOVATime

    Args:
        datalist (dict): the ADO DataList dict, with Key and Value keys

    Returns:
        dict[str, str]: the resulting Pythonic dict
    """
    if 'DataList' in datalist:
        datalist = datalist['DataList']
    data = {}
    for obj in datalist:
        data[obj['Key']] = obj['Value']
    return data


def _build_login_data(secrets, loginpage):
    """
    Build a dict for NOVATime login request

    Args:
        secrets (ConfigParser): the secrets file
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
        'txtUserName': secrets['user']['user'],
        'txtPassword': secrets['user']['password'],
        "hUserAgent": secrets['header']['user-agent'],
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
    return data


def login(secrets):
    """
    Login to NOVATime and return a session

    Args:
        secrets (ConfigParser): the secrets file

    Returns:
        requests.Session: a logged-in session to NOVATime
    """

    # build uri
    cid = secrets['uri']['cid']
    host = secrets['uri']['host']
    page = secrets['uri']['page']
    uri = f"https://{host}/novatime/ewskiosk.aspx"

    # set up headers in secrets and session
    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
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
    secrets.read_dict(
        {'header': headers})
    session = requests.Session()
    session.headers.update(headers)

    # get the login landing page to harvest for login request data
    loginrequest = session.get(uri, params={'CID': cid})
    loginpage = BeautifulSoup(loginrequest.text, 'html.parser')

    # update secrets and session with login cookie parameters
    secrets.read_dict(
        {'cookie': loginrequest.cookies.get_dict(domain=host, path='/')})

    for c, val in secrets['cookie'].items():
        session.cookies.set(c, val, domain=host, path='/')

    # POST the login request to the server
    r = session.post(uri, params={'CID': cid},
                     data=_build_login_data(secrets=secrets, loginpage=loginpage))

    # the SessionVariable contains several pieces of information we need for future requests;
    #  add them to the session headers and the secrets
    user_data_request = session.get(
        url=f'https://{host}/{page}/{cid}/SessionVariable')

    user_data = _ado_to_dict(user_data_request.json())
    if user_data['USERSEQ']:
        user_id = user_data['USERSEQ']
    else:
        user_id = '0'
    employee_id = user_data['EMPSEQ']
    secrets['user']['EmployeeSeq'] = employee_id
    secrets['user']['UserSeq'] = user_id
    session.headers.update(
        {'EmployeeSeq': employee_id,
         'UserSeq': user_id})

    # the employee record contains an access ID we need for other requests, add it to the secrets
    user_data = session.get(
        url=f'https://{host}/{page}/{cid}/employee/{employee_id}').json()

    secrets['user']['AccessSeq'] = str(user_data['Data']['iAccessSeq'])

    return session


if __name__ == '__main__':
    # set up secrets info
    if not isfile('secrets.ini'):
        print('Please copy secrets.ini.example to secrets.ini and configure per the comments', file=sys.stderr)
        sys.exit(-200)
    secrets = configparser.ConfigParser()
    secrets.optionxform = lambda option: option  # return case-sensitive keys
    secrets.read('secrets.ini')

    session = login(secrets)

    # figure out pay period bounds, for URI
    last_pay = arrow.get(
        input('When was your MOST RECENT paycheck, in ISO 8601 (yyyy-mm-dd) format? '))
    start_date, end_date = get_current_pay_period(last_pay)

    # grab current pay period timesheet, write to JSON file in safe directory
    timesheet = get_timesheet(session, secrets, start_date, end_date).json()
    if not isdir('pay'):
        os.mkdir('pay')
    with open(join('pay', f'{start_date.date()}--{end_date.date()}.json'), encoding='utf-8', mode='w') as times:
        json.dump(timesheet, times, indent=' '*4)

    # if the user is not authed, this pukes so handle it gracefully-ish <3
    if 'DataList' not in timesheet:
        print('Not authorized, please check secrets.ini', file=sys.stderr)
        sys.exit(-100)

    # ok we made it! first, get the hours and any exceptions from the current pay period
    hours = get_times(timesheet['DataList'])
    exceptions = get_exceptions(timesheet['DataList'])

    # then, make a little report
    print(f'Pay period from {start_date.date()} to {end_date.date()}:')
    if exceptions:
        print('\tExceptions:')
        for date in exceptions:
            print(f'\t\t{date.format("dddd, MMMM DD, YYYY")}:')
            for exception in exceptions[date]:
                print(f'\t\t\t{exception} = {exceptions[date][exception]}')
    if hours['last_week']:
        print(f'\tLast week: {format_td(hours["last_week"])}')

    # then, more usefully, report how much time left this week,
    #  and, if there is a missing punch today (AKA we are clocked in!), report when to clock out
    #
    # (this if course does no bounds checking or anything so it probably does entertaining things in overtime or Tuesday conditions)
    weekhours = timedelta(hours=int(secrets['hours']['weekhours']))
    remaining = weekhours - hours["this_week"]
    print(
        f'\tThis week: {format_td(hours["this_week"])} ({format_td(remaining)} left)')

    clock_in, clock_out = predict_clock_out(timesheet['DataList'], remaining)

    if clock_out is not None:
        print(
            f'\tAfter clocking in at {clock_in.format("HH:mm")}, clock out by {clock_out.format("HH:mm")} to hit {format_td(weekhours)}')
