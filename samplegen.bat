REM MEMO:Windowsのテスト用バッチファイル
REM MEMO:カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。

REM
REM MEMO:CSVのSimple形式で生成しています。
REM MEMO:ShiftJISで生成しています。UTF-8にする場合は引数 -Cutf-8 を使ってください。
python3 icsconvcsv.py -Cshift_jis 202512 calendar.ics  sjis-202512.csv
python3 icsconvcsv.py -Cshift_jis guess  calendar.ics  sjis-202601.csv
python3 icsconvcsv.py -Cshift_jis all    calendar.ics  sjis-all.csv

REM MEMO:CSVのGaroon形式で生成しています。文字コードはshift_jisです。
python3 ics2gacsv.py 202512 calendar.ics  garoon-202512.csv
python3 ics2gacsv.py guess  calendar.ics  garoon-202601.csv
python3 ics2gacsv.py all    calendar.ics  garoon-all.csv
REM EOF
