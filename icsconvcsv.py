#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025-2026 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import os
import sys
import re
import libicsconvcsv

__doc__=f"""ICS(iCalendar)をCSVに変換。

使用方法:

   $ python3 {sys.argv[0]} [OPTION] 期間 入力.ics 出力.csv

   期間を指定して出力する例:
   $ python3 {sys.argv[0]} 202512 calendar.ics schedules202512.csv
   $ python3 {sys.argv[0]} guess calendar.ics schedules202509.csv
   $ python3 {sys.argv[0]} all calendar.ics schedules-all.csv

   文字コードをShift_JISにする例:
   $ python3 {sys.argv[0]} -Cshift_jis all calendar.ics schedules-sjis.csv
"""
__doc__+=libicsconvcsv.HELP_PART1

__doc__+=libicsconvcsv.HELP_PART2
__doc__+=libicsconvcsv.HELP_LICENSE


########################################

def __myhelp(fname):
    help(fname)
    help("libicsconvcsv")
    sys.exit()

if __name__ == '__main__':
    if libicsconvcsv.VERSION != "3.0":
        print("ERROR: ファイルが古いです。最新のicsconvcsv.pyとlibicsconvcsv.pyをダウンロードしてください。",file=sys.stderr)
        sys.exit(1)

    #CSVを出力する期間指定。
    timerange=0

    exec_filename = os.path.basename(__file__)
    exec_filename = re.sub(r'\.py$', "", exec_filename)

    flag = None

    try:
        argv = sys.argv[1:]

        argv, flag = libicsconvcsv.parse_args(argv, 3)
        if argv is None:
            __myhelp(exec_filename)

        timerange = argv[0]
        input_ics_filename=argv[1]
        output_csv_filename=argv[2]

        timerange = libicsconvcsv.guess_timerange(timerange, input_ics_filename, output_csv_filename)

    except ValueError as e:
        print("ERROR: ", e,  file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit(1)

    libicsconvcsv.ics2csv(flag, input_ics_filename, output_csv_filename, timerange)
#End of main()
