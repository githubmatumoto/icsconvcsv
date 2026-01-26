#!/bin/bash
echo "MEMO: Linuxのテスト用Shell script"
echo "MEMO: カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
echo "MEMO: Excelで読む時は文字コードshift_jisのファイルを使ってください"
#
echo
echo "MEMO: CSVのSimple形式で生成しています。"
echo "MEMO: 文字コードutf_8で生成"
python3 icsconvcsv.py 202512 calendar.ics  utf8-202512.csv
python3 icsconvcsv.py guess  calendar.ics  utf8-202601.csv
python3 icsconvcsv.py  all   calendar.ics  utf8-all.csv
echo
echo "MEMO: shift_jis で生成しています。"
python3 icsconvcsv.py -Cshift_jis 202512 calendar.ics  sjis-202512.csv
python3 icsconvcsv.py -Cshift_jis guess  calendar.ics  sjis-202601.csv
python3 icsconvcsv.py -Cshift_jis all    calendar.ics  sjis-all.csv
echo
echo "MEMO: CSVのGaroon形式で生成しています。文字コードはshift_jisです。"
python3 ics2gacsv.py 202512 calendar.ics  garoon-202512.csv
python3 ics2gacsv.py guess  calendar.ics  garoon-202601.csv
python3 ics2gacsv.py all    calendar.ics  garoon-all.csv
#EOF
