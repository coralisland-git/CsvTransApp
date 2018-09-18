import re
import sys
import numbers


def validate_phone_number(val, row, row_num, col_name, quit_on_error):
    if not val:
        return ''

    if isinstance(val, float):
        val = int(val)

    val = str(val)

    for c in ['(', ')', '/', ' ', '-']:
        val = val.replace(c, '')

    if re.compile("^[0-9]{10}$").match(val):
        return val[:3] + '-' + val[3:6] + '-' + val[6:10]
    else:
        return ''


def validate_ssn(val, row, row_num, col_name, quit_on_error):
    if not val:
        return ''

    if isinstance(val, float):
        val = int(val)

    val = ''.join(re.findall(r'\d+', str(val)))

    if re.compile("^[0-9]{9}$").match(val):
        return val[:3] + '-' + val[3:5] + '-' + val[5:9]
    else:
        return ''


def validate_number(val, row, row_num, col_name, quit_on_error):
    if not val:
        return ''

    if isinstance(val, numbers.Number):
        return val

    if isinstance(val, str):
        try:
            return float(val)
        except Exception:
            pass

    print('Error: {} did not contain a numerical value. Value: "{}"'.format(col_name, val))
    print('Row - ', dict(row))
    sys.exit(-1)

