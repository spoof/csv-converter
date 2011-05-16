#!/usr/bin/env python
# encoding: utf-8
"""
csv_converter.py

Created by Sergey Safonov <spoof@spoofa.info> on 2011-05-16.
Copyright (c) 2011 . All rights reserved.
"""

import sys
import xlwt
import csv
import cStringIO
import getopt


help_message = '''Usage: %s [OPTION]... infile.csv outfile.xls
Converts CSV file to XLS
Options:
  -d\tdelimeter used in CSV file''' % sys.argv[0]


class CsvConverter(object):
    def __init__(self, f, delimiter=' '):
        try:
            self.f = open(f, 'rb')
        except IOError, e:
            print >> sys.stderr, "Cannot open CSV file:", str(e)
            raise

        self.delimiter = delimiter

    def convert(self):
        """Converts contents of CSV file to XLS
        """
        reader = csv.reader(self.f, delimiter=self.delimiter, quotechar='|')
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet 1')
        xls = cStringIO.StringIO()

        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                ws.write(r, c, col)

        wb.save(xls)
        return xls.getvalue()

if __name__ == "__main__":
    delimiter = " "
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hd:", ["help"])
    except getopt.GetoptError, err:
        print str(err)
        sys.exit(2)

    for o, a in opts:
        if o in ("-h", "--help"):
            print help_message
            sys.exit()
        elif o == "-d":
            delimiter = a
    if len(args) < 2:
        print help_message
        sys.exit(2)

    xlsfile = args[1]
    try:
        c = CsvConverter(args[0], delimiter=delimiter)
    except IOError:
        sys.exit(2)

    output = c.convert()
    try:
        xlsfile = open(xlsfile, 'w')
        xlsfile.write(output)
    except IOError, e:
        print >> sys.stderr, "Error while writing XLS file:", e
        sys.exit(2)

    sys.exit(0)
