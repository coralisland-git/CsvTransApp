import re
import sys
import numbers
import decimal

def validate_phone_number(d, row, row_num):
    if not d:
        return ''

    d = str(d)

    for c in ['(', ')', '/', ' ', '-']:
        d = d.replace(c, '')

    if re.compile("^[0-9]{3}[0-9]{3}[0-9]{4}$").match(d):
        return d[:3] + '-' + d[3:6] + '-' + d[6:10]
    else:
        return ''

def validate_payroll_frequency(d, row, row_num):
    if isinstance(d, numbers.Number):
        return d

    if isinstance(d, str):
        try:
            return float(d)
        except Exception:
            pass

    print('Error on row {} - Payroll Frequency did not contain a numerical value'.format(row_num + 1))
    print('Found: {}'.format(d))
    sys.exit(-1)

def set_hire_date(d, row, row_num):
    rehire_date = row[23]
    if rehire_date and rehire_date > d:
            return rehire_date
    return d
