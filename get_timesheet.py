# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL

"""Get a NOVATime timesheet and print a little report"""








def get_current_pay_period(last_pay):
    """Calculate the bounds of the pay period.

    Our pay periods are two weeks long, ending one week after a paycheck:
        Always begins on a Monday (weekday = 0)
        Always ends on a Sunday (weekday = 6)

    Args:
        last_pay (arrow): the last paycheck

    Returns a tuple with:
        start_date (arrow): the beginning of the pay period
        end_date (arrow): the end of the pay period

    """
    end_date = last_pay.shift(weeks=1, weekday=6)

    # pay period starts two weeks before the end, on a Monday (=0)
    start_date = end_date.shift(weeks=-2, weekday=0)
    return (start_date, end_date)





def main():
    """Main function"""
    

    session = login(secrets)

    # figure out pay period bounds, for URI
    last_pay = arrow.get(
        input('When was your MOST RECENT paycheck, in ISO 8601 (yyyy-mm-dd) format? '))
    start_date, end_date = get_current_pay_period(last_pay)

    # grab current pay period timesheet, write to JSON file in safe directory
    timesheet = get_timesheet(session, secrets, start_date, end_date)

    


if __name__ == '__main__':
    main()
