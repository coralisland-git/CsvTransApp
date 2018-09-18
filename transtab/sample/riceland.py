
def set_hire_date(row, row_num, quit_on_error):
    if not row['Hire Date']:
        row['Hire Date'] = row['Job Rehire Date']
    elif row['Job Rehire Date']:
        if row['Job Rehire Date'] > row['Hire Date']:
            row['Hire Date'] = row['Job Rehire Date']
    return row
