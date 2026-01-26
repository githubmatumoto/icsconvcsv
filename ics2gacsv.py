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

__doc__ = f"""ICS(iCalendar)をCSVに変換。CSVの出力形式はGaroonとほぼ同じ。

使用方法:
   $ python3 {sys.argv[0]} [-hsmz] 期間 入力.ics 出力.csv

   期間を指定して出力する例:
   $ python3 {sys.argv[0]} 202512 calendar.ics schedules202512.csv
   $ python3 {sys.argv[0]} guess calendar.ics schedules202509.csv
   $ python3 {sys.argv[0]} all calendar.ics schedules-all.csv
"""

__doc__ += libicsconvcsv.HELP_PART1

__doc__ += """
=====================================================================
注意事項:
   ※過去に公開していたics2gacsvの互換コマンドになります。
   ※大部分のオプションは無効になっていますので、細かい指定
     を行う場合はicsconvcsv.pyを直接利用ください。

   現在の仕様としては、下記引数で起動した場合と同等です。

    $ python3 icsconvcsv.py -FGaroon 期間 入力.ics 出力.csv

以上です。
"""
__doc__ += libicsconvcsv.HELP_LICENSE


########################################

def __myhelp(fname):
    help(fname)
    sys.exit()

if __name__ == '__main__':
    if libicsconvcsv.VERSION != "3.0":
        print("ERROR: ファイルが古いです。最新のics2gacsv.pyとlibicsconvcsv.pyをダウンロードしてください。", file=sys.stderr)
        sys.exit(1)

    #CSVを出力する期間指定。
    timerange = 0

    exec_filename = os.path.basename(__file__)
    exec_filename = re.sub(r'\.py$', "", exec_filename)

    # 許可する引数。
    short_opt = 'hsmzk'
    long_opt = ["help", "format-garoon", "disable-split-summary",\
                "extend-summary-head", "add-summary-head=",\
                "enhance-gyoumunum", "enhance-gyoumu-number",\
                "print-csv-header", "DEBUG-UID="]

    # ライブラリの挙動変更。
    ext_argv = ['--format-garoon']

    flag = None

    try:
        argv = ext_argv + sys.argv[1:]

        argv, flag = libicsconvcsv.parse_args(argv, 3, short_opt, long_opt)
        if argv is None:
            __myhelp(exec_filename)

        timerange = argv[0]
        input_ics_filename = argv[1]
        output_csv_filename = argv[2]

        timerange = libicsconvcsv.guess_timerange(timerange, input_ics_filename, output_csv_filename)

    except ValueError as e:
        print("ERROR: ", e, file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit(1)

    libicsconvcsv.ics2csv(flag, input_ics_filename, output_csv_filename, timerange)
#End of main()
