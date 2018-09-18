from optparse import OptionParser

def process_options():
    parser = OptionParser()
    parser.add_option("-f", "--format", dest="out_format",help="enter format", metavar="FORMAT")
    parser.add_option("-i", "--input", dest="in_f", help="input file path", metavar="PATH")
    parser.add_option("-s", "--sheet", dest="out_sheet", help="sheet name", metavar="SHEET_NAME")

    (options, args) = parser.parse_args()

    if not options.in_f:
        parser.error('input file not provided (-i option)')

    options.in_f = options.in_f.rstrip('/').rstrip('\\')

    if not options.out_format:
        parser.error('format not provided (-f option)')

    return options.in_f, options.out_format, options.out_sheet