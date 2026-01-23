#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2026 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import sys
import csv
import re
import getopt

__doc__="""
CSV比較用。

STDINから読み込んだCSVの各要素を文字列ソートして出力する。

一行めに(N/A)があれば空文字にする。

末尾の改行空白などを除去する。
空行を除去する。
"""

argv = sys.argv[1:]

opts, argv = getopt.gnu_getopt(argv, "12345a")

print_line=1000
for o, a in opts:
    if o == "-1":
        print_line = 1
    elif o  == "-2":
        print_line = 2
    elif o  == "-3":
        print_line = 3
    elif o  == "-4":
        print_line = 4
    elif o  == "-5":
        print_line = 5
    elif o == "-a":
        print_line=1000

reader = csv.reader(sys.stdin)

csv_buffer = []
for row in reader:
    new_row = []
    for i in row:
        if re.search(r"^\(N\/A\)", i, flags=re.DOTALL):
            new_row.append("")
            continue

        lines = i.splitlines()
        
        n = []
        for j in range(len(lines)):
            if j >= print_line:
                break
            n.append(lines[j].rstrip())


        j = ("\n".join(n)).rstrip()
        new_row.append(j)
    csv_buffer.append(new_row)

#文字列ソート
csv_buffer.sort()

#CSV Open
csv_writer = csv.writer(sys.stdout,quoting=csv.QUOTE_ALL)

for i in csv_buffer:
    csv_writer.writerow(i)
#EOF
