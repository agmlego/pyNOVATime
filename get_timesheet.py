# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL

import arrow
import configparser
import json
import sys
import os
from os.path import isdir,isfile,join
from datetime import timedelta
import requests

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
    end_date = last_pay.shift(weeks=1,weekday=6)

    # pay period starts two weeks before the end, on a Monday (=0)
    start_date = end_date.shift(weeks=-2,weekday=0)
    return (start_date,end_date)

def get_timesheet(secrets,start_date,end_date):
    '''Log in and download the timesheet.

    Here there be dragons. This does no error-checking, and is brittle as most scrapers are.

    Keyword arguments:
        secrets -- a `ConfigParser` for the secrets file (see `secrets.ini.example`)
        start_date -- an `arrow` indicating the beginning of the pay period
        end_date -- an `arrow` indicating the end of the pay period

    Returns a `requests.models.Response` from the site, hopefully containing the timesheet as JSON

    '''
    user = secrets['user']['user']
    password = secrets['user']['password']

    # build uri, parameters, and headers
    cid = secrets['site']['UsedNOVA4000CID']
    host = secrets['uri']['host']
    page = secrets['uri']['page']

    uri = f"https://{host}/{page}/{cid}/timesheetdetail"
    parameters = {
        'AccessSeq': secrets['user']['accessseq'],
        'EmployeeSeq': secrets['user']['employeeseq'],
        'StartDate': start_date.format(PARAM_DATE_FORMAT),
        'EndDate': end_date.format(PARAM_DATE_FORMAT),
        'UserSeq': secrets['user']['userseq'],
        'CustomDateRange': False,
        'ShowOneMoreDay': False,
        'EmployeeSeqList': '',
        'DailyDate': start_date.format(PARAM_DATE_FORMAT),
        'ForceAbsent': False,
        'PolicyGroup': ''
    }
    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'en-US,en;q=0.5',
        'connection': 'keep-alive',
        'cookie': 'MyTZ=VALUE=1013; MyDLSavings=VALUE=True; UsedNOVA4000ThemePath=; '
            f"UsedNOVA4000CID={cid}; "
            'BackdropImageURL=https://online.timeanywhere.com/wp/landscape-403165_1920.jpg; '
            f"NOVA_cookie-47873={secrets['cookie']['NOVA_cookie-47873']}; i18next=en; defaultLocale=en-US; "
            f"UsedNOVA4000ClientID={secrets['site']['UsedNOVA4000ClientID']}; "
            f"ASP.NET_SessionId={secrets['cookie']['ASP.NET_SessionId']}",
        'dnt': '1',
        'employeeseq': secrets['user']['employeeseq'],
        'locale': 'en-US',
        'referer': f'https://{host}/TimeanywhereExt/loadTimesheet.aspx',
        'user-agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:68.0) Gecko/20100101 Firefox/68.0',
        'userseq': secrets['user']['userseq']
    }

    return requests.get(uri,headers=headers,auth=(user,password),params=parameters)

def parse_punch(punch_str):
    '''Parse a punch entry into an `arrow`'''
    return arrow.get(punch_str,PUNCH_FORMAT).replace(tzinfo='America/Detroit')

def parse_date(date_str):
    '''Parse a date entry into an `arrow`'''
    return arrow.get(date_str,NOVA_DATE_FORMAT).replace(tzinfo='America/Detroit')

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
        exceptions[punch_date] = {key:entry[key] for key in entry if 'Exception' in key and entry[key]}
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

def predict_clock_out(timesheet,remaining):
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
        return None,None
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
    return clock_in,clock_out


if __name__ == '__main__':
    # set up secrets info
    if not isfile('secrets.ini'):
        print('Please copy secrets.ini.example to secrets.ini and configure per the comments',file=sys.stderr)
        sys.exit(-200)
    secrets = configparser.ConfigParser()
    secrets.read('secrets.ini')

    # figure out pay period bounds, for URI
    last_pay = arrow.get(input('When was your MOST RECENT paycheck, in ISO 8601 (yyyy-mm-dd) format? '))
    start_date,end_date = get_current_pay_period(last_pay)
    
    # grab current pay period timesheet, write to JSON file in safe directory
    timesheet = get_timesheet(secrets,start_date,end_date).json()
    if not isdir('pay'):
        os.mkdir('pay')
    with open(join('pay',f'{start_date.date()}--{end_date.date()}.json'), encoding='utf-8', mode='w') as times:
        json.dump(timesheet,times,indent=' '*4)

    # if the user is not authed, this pukes so handle it gracefully-ish <3
    if 'DataList' not in timesheet:
        print('Not authorized, please check secrets.ini',file=sys.stderr)
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
    print(f'\tThis week: {format_td(hours["this_week"])} ({format_td(remaining)} left)')

    clock_in,clock_out = predict_clock_out(timesheet['DataList'],remaining)

    if clock_out is not None:
        print(f'\tAfter clocking in at {clock_in.format("HH:mm")}, clock out by {clock_out.format("HH:mm")} to hit {format_td(weekhours)}')