import os
import sys
import maya
import importlib
from tablib import Dataset, Databook
from warnings import filterwarnings
filterwarnings("ignore", category=UserWarning)

from pyparsing import *

from xlrd import xldate
from xlrd.sheet import ctype_text

from transtab.utils import process_options
from transtab.global_operations import *

from transtab.tablib_patch import dset_sheet as dset_sheet_patched
import tablib
tablib.formats.xlsx.dset_sheet = dset_sheet_patched


###############################################################################
#                         Grammar for format specification                    #
###############################################################################
LPARA, RPARA, COLON, COMMA, EQUALS = map(Suppress, "{}:,=")

DATES = CaselessLiteral("dates").setResultsName('decl')

NEW = CaselessLiteral("new").setResultsName('op')
CLEAR = CaselessLiteral("clear").setResultsName('op')
DELETE = CaselessLiteral("delete").setResultsName('op')
RENAME = CaselessLiteral("rename").setResultsName('op')
COPY = CaselessLiteral("copy").setResultsName('op')
CUTPASTE = CaselessLiteral("cutpaste").setResultsName('op')
CONCATENATE = CaselessLiteral("concatenate").setResultsName('op')
REPLACE = CaselessLiteral("replace").setResultsName('op')
DELETE_DUPLICATE_ROWS = CaselessLiteral("delete-duplicate-rows").setResultsName('op')
DELETE_ROWS_BY_COLUMN_VAL = CaselessLiteral("delete-rows-by-column-val").setResultsName('op')
DROP = CaselessLiteral("drop").setResultsName('op')
SUM_COL_AND_DELETE_DUPLICATE_ROWS = CaselessLiteral("sum-col-and-delete-duplicate-rows").setResultsName('op')
DO = CaselessLiteral("do").setResultsName('op')
CUSTOM_OP_NAME = Word( alphas+"_", alphanums+"_" ).setResultsName('custom_op_name')

COL = CaselessLiteral("col").setResultsName('col')
ROW = CaselessLiteral("row").setResultsName('row')


IN = Suppress(CaselessLiteral("in"))
ON = Suppress(CaselessLiteral("on"))
AS = Suppress(CaselessLiteral("as"))
TO = Suppress(CaselessLiteral("to"))
AND = Suppress(CaselessLiteral("and"))
USING = Suppress(CaselessLiteral("using"))
STORE = Suppress(CaselessLiteral("store"))
DEFAULT = CaselessLiteral("default")
CASE_INSENSITIVE = CaselessLiteral("case-insensitive")
QUIT_ON_ERROR = CaselessLiteral("quit-on-error")
UNIQUE = Suppress(CaselessLiteral("unique"))
SUM = Suppress(CaselessLiteral("sum"))
DEL_COL = Suppress(CaselessLiteral("col"))
DEL_VAL = Suppress(CaselessLiteral("val"))

ROW_NUM = Word(nums).setResultsName('row_num').setParseAction( lambda t: int(t[0]) )

COL_NAME = QuotedString("'").setResultsName('col_name')

COL_LIST = Group(delimitedList(COL_NAME))

KEY = QuotedString("'").setResultsName('k') + COLON
VAL = QuotedString("'").setResultsName('v') + Optional(COMMA)
KV_MAP =  LPARA + dictOf(KEY, VAL).setResultsName('kv_map') + RPARA

REPLACE_DEFAULT_PARAM = DEFAULT.setResultsName('has_default') + QuotedString("'").setResultsName('default_val')
REPLACE_CASE_FLAG = CASE_INSENSITIVE.setResultsName('case_insensitive')
REPLACE_PARAMS = OneOrMore(REPLACE_DEFAULT_PARAM | REPLACE_CASE_FLAG) 

DATES_DECL = DATES + EQUALS + COL_LIST.setResultsName('cols')

NEW_CMD = NEW + Optional(COL) + COL_NAME
CLEAR_CMD = CLEAR + COL_NAME
DELETE_CMD = (DELETE + ROW + ROW_NUM) | (DELETE + COL_NAME)
RENAME_CMD = RENAME + COL_NAME.setResultsName('col_name') + Optional(AS) + COL_NAME.setResultsName('new_name')
COPY_CMD = COPY + COL_NAME.setResultsName('src_col') + Optional(TO) + COL_NAME.setResultsName('dest_col')
CUTPASTE_CMD = CUTPASTE + COL_NAME.setResultsName('src_col') + Optional(TO) + COL_NAME.setResultsName('dest_col')
CONCATENATE_CMD = CONCATENATE + COL_LIST.setResultsName('src_cols') + Optional(AND) + Optional(STORE) + Optional(IN) + COL_NAME.setResultsName('dest_col') + Optional(USING + QuotedString("'").setResultsName('join_str'))
REPLACE_COL_CMD = REPLACE + COL_NAME + KV_MAP + Optional(REPLACE_PARAMS)
DELETE_DUPLICATES_CMD = DELETE_DUPLICATE_ROWS + Optional(UNIQUE + COL_NAME)
DELETE_ROWS_BY_COLUMN_VAL_CMD = DELETE_ROWS_BY_COLUMN_VAL + DEL_COL + COL_NAME.setResultsName('col') + DEL_VAL + COL_NAME.setResultsName('val')
DROP_CMD = DROP
SUM_DELETE_DUPLICATES_CMD = SUM_COL_AND_DELETE_DUPLICATE_ROWS + SUM + COL_NAME.setResultsName('sum_col') + UNIQUE + COL_NAME.setResultsName('unique_col')
CUSTOM_CMD = DO + CUSTOM_OP_NAME + Optional(ON + COL_NAME) + Optional(QUIT_ON_ERROR)

F_CMD = Group( DATES_DECL
    | NEW_CMD 
    | CLEAR_CMD 
    | DELETE_CMD 
    | RENAME_CMD 
    | COPY_CMD
    | CUTPASTE_CMD 
    | CONCATENATE_CMD 
    | REPLACE_COL_CMD 
    | DELETE_DUPLICATES_CMD 
    | DROP_CMD
    | SUM_DELETE_DUPLICATES_CMD 
    | CUSTOM_CMD 
    | DELETE_ROWS_BY_COLUMN_VAL_CMD)

COMMENT = Suppress(pythonStyleComment)

GRAMMAR = OneOrMore(F_CMD + Optional(COMMENT) | COMMENT)


###############################################################################
#                                Main class                                   #
###############################################################################

class TransTab(object):

    def __init__(self, in_fname, format_fname, out_sheet='Sheet1', out_fname=''):

        self.in_fname = in_fname

        self.in_fname_prefix, self.in_type = os.path.splitext(in_fname)
        self.in_type = self.in_type.lstrip('.')

        self.out_fname_prefix, self.out_type = os.path.splitext(out_fname)
        self.out_type = self.out_type.lstrip('.')

        if not self.out_fname_prefix or not self.out_type:
            self.out_fname_prefix = self.in_fname_prefix + '_formatted'
            self.out_type = 'xlsx'

        self.out_sheet = out_sheet

        with open(format_fname, 'r') as f:
            self.file_format = f.read()

        sys.path.append(os.path.dirname(os.path.abspath(format_fname)))

        self.format_name = os.path.splitext(os.path.basename(format_fname))[0]

        try:
            self.format_module = importlib.import_module(self.format_name)
        except:
            self.format_module = ''

        if self.in_type in ['xls', 'xlsx']:
            self.file_read_mode = 'rb'
        else:
            self.file_read_mode = 'r'

        with open(in_fname, self.file_read_mode) as f:
            self.data = Dataset().load(f.read())

        if self.in_type in ['xls', 'xlsx']:
            for i, row in enumerate(self.data):
                self.data[i] = [int(n) if type(n)==float and n == int(n) else n for n in row]


    def preprocess_dates(self, cols):
        for i, row in enumerate(self.data.dict):
            for col in cols:
                if not row[col]:
                    continue

                try:
                    if self.in_type in ['xls', 'xlsx']:
                        row[col] = xldate.xldate_as_datetime(row[col], 0)
                    else:
                        row[col] = maya.when(row[col]).datetime().replace(tzinfo=None)
                except:
                    row[col] = maya.when(row[col]).datetime().replace(tzinfo=None)

                self.data[i] = list(row.values())


    def new_col(self, col_name):
        self.data.append_col(lambda x: '', header=col_name)


    def clear_col(self, col_name):
        pos = self.data.headers.index(col_name)
        self.delete_col(col_name)
        self.data.insert_col(pos, lambda x: '', header=col_name)

    def delete_rows_by_column_val(self, col_name, val):
        arr = self.data[col_name]
        removed_count = 0
        for ind in range(0, len(arr)):
            if val == '':
                if arr[ind] == val:
                    del self.data[ind-removed_count]
                    removed_count += 1    
            else:
                if val in arr[ind]:
                    del self.data[ind-removed_count]
                    removed_count += 1

    def delete_row(self, n):
        n = n - 2
        if n == -1:
            new_data = Dataset()
            new_data.headers = self.data[0]
            for row in self.data[1:]:
                new_data.append(row)
            self.data = new_data
        else:
            del self.data[n]


    def delete_col(self, col_name):
        del self.data[col_name]


    def drop(self):
        n = len(self.data) - 1
        if n == -1:
            new_data = Dataset()
            new_data.headers = self.data[0]
            for row in self.data[1:]:
                new_data.append(row)
            self.data = new_data
        else:
            del self.data[n]

    def rename_col(self, col_name, new_name):
        pos = self.data.headers.index(col_name)
        self.data.headers[pos] = new_name


    def copy_col(self, src_col, dest_col):
        if dest_col not in self.data.headers:
            self.new_col(dest_col)

        src_pos = self.data.headers.index(src_col)
        dest_pos = self.data.headers.index(dest_col)
        self.delete_col(dest_col)       
        self.data.insert_col(dest_pos, lambda row: row[src_pos], header=dest_col)


    def cutpaste_col(self, src_col, dest_col):
        if dest_col not in self.data.headers:
            self.new_col(dest_col)

        self.copy_col(src_col, dest_col)
        self.clear_col(src_col)


    def concatenate_col(self, src_cols, dest_col, join_str):
        if dest_col not in self.data.headers:
            self.new_col(dest_col)

        for i, row in enumerate(self.data.dict):
            row[dest_col] = join_str.join([str(row[c]) for c in src_cols])
            self.data[i] = list(row.values())


    def replace(self, col_name, kv_map, has_default, default_val, case_insensitive):
        pos = self.data.headers.index(col_name)

        if case_insensitive:
            val = (i.lower() for i in self.data[col_name])
            kv_map = {k.lower(): v for k, v in kv_map.items()}
        else:
            val = (i for i in self.data[col_name])

        self.delete_col(col_name)
        self.data.insert_col(pos, lambda row: kv_map.get(str(next(val)), default_val) if has_default else kv_map[str(next(val))], header=col_name)


    def delete_duplicates(self, col_name=''):
        if not col_name:
            self.data.remove_duplicates()
        else:
            seen = set()
            pos = self.data.headers.index(col_name)
            self.data._data[:] = [row for row in self.data._data if not (row[pos] in seen or seen.add(row[pos]))]


    def sum_delete_duplicates(self, sum_col, unique_col):
        unique_val_sum = {}
        unique_val_row1 = {}

        # Collect sum of sum_cols per unique_col and also the first row number per unique col value
        for i, row in enumerate(self.data.dict):
            sum_col_val = row[sum_col]
            unique_col_val = row[unique_col]

            if unique_col_val in unique_val_sum:
                unique_val_sum[unique_col_val] += sum_col_val
            else:
                unique_val_sum[unique_col_val] = sum_col_val
                unique_val_row1[unique_col_val] = i

        # Set the sum for each unique col value in the respective first occurrence rows
        for unique_col_val in unique_val_row1:
            first_row_num = unique_val_row1[unique_col_val]
            row = self.data.dict[first_row_num]
            row[sum_col] = unique_val_sum[unique_col_val]
            self.data[first_row_num] = list(row.values())

        self.delete_duplicates(unique_col)


    def get_custom_func(self, name):
        '''
        Suppose the format is specified as <_some_path_/riceland.txt>.
        The program tries to load the custom operations definition file at <_some_path_/riceland.py>.

        There is also a global_operations.py file that contains custom operations definitions.
        If you modify this, then the program has to be reinstalled for it to be available at commandline.

        '''
        if hasattr(self.format_module, name):
            func_obj = getattr(self.format_module, name)
            if name in globals():
                print('{} defined both globally and for the format. Using the format definition'.format(name))
        else:
            if name in globals():                   
                func_obj = globals()[name]
            else:
                print('{} is not defined'.format(name))
                sys.exit(-1)
        return func_obj


    def do_custom_operation(self, op, quit_on_error=False):
        '''
        This operates on a single column. 
        Done when there is an instruction of the kind: do <custom_op_name> on <col_name>

        The custom function executed by passing the following parameters:
        the row dict, row number, flag whether to quit if there is an error.

        It expects the modified row dict as the return value which will be assigned back to the main data.
        '''
        for i, row in enumerate(self.data.dict):
            self.data[i] = list(op(row, i, quit_on_error).values())


    def do_custom_operation_col(self, op, col_name, quit_on_error=False):
        '''
        This operates on a single column. 
        Done when there is an instruction of the kind: do <custom_op_name> on <col_name>

        The custom function executed by passing the following parameters:
        the cell value, row dict, row number, column number, flag whether to quit if there is an error.

        It expects the modified vell value as the return value which will be assigned back to the column in main data.
        '''
        for i, row in enumerate(self.data.dict):
            row[col_name] = op(row[col_name], row, i, col_name, quit_on_error)
            self.data[i] = list(row.values())


    def transform(self):
        for f_cmd in GRAMMAR.parseString(self.file_format):
            if f_cmd.decl == 'dates':
                self.preprocess_dates(f_cmd.cols)

            elif f_cmd.op == 'new':
                self.new_col(f_cmd.col_name)

            elif f_cmd.op == 'clear':
                self.clear_col(f_cmd.col_name)

            elif f_cmd.op == 'delete':
                if f_cmd.row:
                    self.delete_row(f_cmd.row_num)
                else:
                    self.delete_col(f_cmd.col_name)

            elif f_cmd.op == 'drop':
                self.drop()

            elif f_cmd.op == 'rename':
                self.rename_col(f_cmd.col_name, f_cmd.new_name)

            elif f_cmd.op == 'copy':
                self.copy_col(f_cmd.src_col, f_cmd.dest_col)

            elif f_cmd.op == 'cutpaste':
                self.cutpaste_col(f_cmd.src_col, f_cmd.dest_col)

            elif f_cmd.op == 'concatenate':
                self.concatenate_col(f_cmd.src_cols, f_cmd.dest_col, f_cmd.join_str)

            elif f_cmd.op == 'replace':
                self.replace(f_cmd.col_name, f_cmd.kv_map, f_cmd.has_default, f_cmd.default_val, f_cmd.case_insensitive)

            elif f_cmd.op == 'delete-duplicate-rows':
                self.delete_duplicates(f_cmd.col_name)

            elif f_cmd.op == 'delete-rows-by-column-val':
                self.delete_rows_by_column_val(f_cmd.col, f_cmd.val)

            elif f_cmd.op == 'sum-col-and-delete-duplicate-rows':
                self.sum_delete_duplicates(f_cmd.sum_col, f_cmd.unique_col)

            elif f_cmd.op == 'do' and f_cmd.custom_op_name:
                func_obj = self.get_custom_func(f_cmd.custom_op_name)

                if f_cmd.col_name:
                    self.do_custom_operation_col(func_obj, f_cmd.col_name, f_cmd.quit_on_error)
                else:
                    self.do_custom_operation(func_obj, f_cmd.quit_on_error)

            else:
                print("Don't know how to process this instruction: {}".format(f_cmd))

        self.save()


    def save(self):
        book = Databook()
        self.data.title = self.out_sheet
        book.add_sheet(self.data)
        with open(self.out_fname_prefix + '.' + self.out_type, 'wb') as f:
            f.write(book.export(self.out_type))


###############################################################################
#            To use when Transtab is used by executing this very file         #
###############################################################################

def main():
    in_f, format_fname, out_sheet = process_options()
    TransTab(in_fname = in_f, format_fname = format_fname, out_sheet=out_sheet).transform()

if __name__ == '__main__':
    main()
