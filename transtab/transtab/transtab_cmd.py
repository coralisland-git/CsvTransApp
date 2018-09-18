import sys

from transtab.TransTab import TransTab
from transtab.utils import process_options

def main():
    in_f, format_fname, out_sheet = process_options()

    tt = TransTab(in_fname = in_f, format_fname = format_fname, out_sheet=out_sheet)
    tt.transform()

if __name__ == '__main__':
    main()
    
