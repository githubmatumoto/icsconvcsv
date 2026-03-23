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
import datetime
import libicsconvcsv

# 入力ファイル名
#Outlook(Web)が生成するICSファイルの出力名
__INPUT_ICS_FILENAME = "calendar.ics"

__doc__ = f"""ICS(iCalendar)をCSVに変換。作者の職場の業務記録提出用。

SYNOPSIS:

  $ python3 {sys.argv[0]} NAME

  引数: NAME: 業務記録提出者の氏名。

DESCRIPTION:

ソフトウエアicsconvcsvに同梱されるスクリプトの一つです。作者の職場の業
務記録提出用のスクリプトです。入力/出力ファイル名が決め打ちになってい
ます。

業務記録の記述方法は2026年3月に改訂されています。また出力ファイル形式
がCSVからICSに変更になっています。

本スクリプトは2026年3月以降の記述方法の業務記録を古い記述方法のCSVファ
イルに変換します。

本プログラムを初めて利用する場合は、README.md および INSTALL.mdをエディ
タで参照して、ソフトウエアicsconvcsvの初期設定を行ってください。

以下は初期設定を行った後の手順書になります。

1. ソフトウエアicsconvcsvを展開したフォルダにICSファイルを置く。
ファイル名は「{__INPUT_ICS_FILENAME}」で置いてください。

ICSファイルの取扱いには注意ください。特にTeams会議のパスワードが含まれ
ます。

ICSファイルはOutlook(Web)からダウンロードするのを勧めます。
Outlook(classic)からダウンロードすると、Teams会議参加者のメールアドレ
スが含まれます。漏洩すると個人情報の流出となります。

2. コマンドプロンプトで毎回必要な初期設定や確認事項の確認を行う。

Linux/macOSは以下を実行してください。Pythonの初期化になります。

  $ source ~/.icsconvcsv/bin/activate

Windowsはライブラリvobjectを導入したpythonと同じであるか確認する。

  > python3 --version

3. コマンドプロンプトでソフトウエアを展開したフォルダに移動する。

$ cd "ソフトウエアを展開したフォルダ"

4. 下記コマンドを実行する。NAMEには業務記録提出者の氏名を記載する

  $ python3 {sys.argv[0]} NAME

成功すると、CSV形式の業務記録が2個生成されます。実行した当月と、その前
の月になります。

例: 引数NAMEに工大太郎を指定して2026年1月に実行すると、出力ファイル名
は以下になります。

  schedules202512工大太郎.csv
  schedules202601工大太郎.csv

5. CSVをExcelで確認し、個人のプライベートスケジュールが含まれてないか
確認し、必要に応じて削除を行ってください。

注意事項:

※同梱されている「icsconvcsv.py」の簡易版になります。現在の仕様として
は、下記引数で起動した場合と同等です。

  $ python3 icsconvcsv.py --enable-file-exist-test -k -m (継続行)
        --convert-old-style-tourokunum (継続行)
        --delete-location=MicrosoftTeams会議:MicrosoftTeamsMeeting (継続行)
        -FGaroon -Esimple guess {__INPUT_ICS_FILENAME} schedules今月NAME.csv

  $ python3 icsconvcsv.py --enable-file-exist-test -k -m (継続行)
        --convert-old-style-tourokunum (継続行)
        --delete-location=MicrosoftTeams会議:MicrosoftTeamsMeeting (継続行)
        -FGaroon -Esimple guess {__INPUT_ICS_FILENAME} schedules先月NAME.csv

  "今月"と"先月"は20yynnの6桁の数字

※作者の職場の業務記録提出用のため頻繁に仕様が変更になります。
"""

__doc__ += libicsconvcsv.HELP_LICENSE

########################################

def __myhelp(fname):
    help(fname)
    sys.exit()

if __name__ == '__main__':
    if libicsconvcsv.VERSION != "3.2":
        print("ERROR: ファイルが古いです。最新のkiroku.pyとlibicsconvcsv.pyをダウンロードしてください。", file=sys.stderr)
        sys.exit(1)

    #####################################################################
    # key: 出力期間, value:出力ファイル名
    csv_fname_list = {}

    #####################################################################
    # 引数解析
    exec_filename = os.path.basename(__file__)
    exec_filename = re.sub(r'\.py$', "", exec_filename)

    # 許可する引数。
    # short_optから引数zを削除(2026/3/23)
    short_opt = 'hWkmE:C:'
    long_opt = ["format-garoon"]
    long_opt += ["help", "enable-file-exist-test", "add-summary-head="]

    # 本採用の業務記録簿の仕様追加(2026/3/23)
    long_opt += ["delete-location=", 'convert-old-style-tourokunum']

    # ライブラリの挙動変更。

    # CSV出力はGaroon Format指定
    ext_argv = ['--format-garoon']

    #※本仕様は不採用となったため、保守していません(2026/3/23)
    # 拡張登録番号を指定。
    #ext_argv += ['-z']

    # CSVの１行めに項目(CSVヘッダ)を出力
    ext_argv += ['-k']
    # SUMMARYヘッダ拡張。
    ext_argv += ['-m']

    # 煩雑なファイル存在確認を停止したい場合は以下をコメントアウト
    ext_argv += ['--enable-file-exist-test']
    #
    # SUMMARYヘッダに独自の拡張を行いたい場合は下記のような形で追記ください。
    # ext_argv += ['--add-summary-head=研究,教育']

    # UTF-8からUTF-8以外に変換時、ギリシャ数字や丸付き数字などをアスキー文字に変換します。
    # この処理により情報量が落ちます。停止したい場合は以下をコメントアウト
    ext_argv += ['-Esimple']

    # 2026年3月改訂の業務記録の仕様に準じて追加(2026/3/23)
    # 指定したLOCATONが含まれるVEVENTを削除。
    ext_argv += ['--delete-location=MicrosoftTeams会議:MicrosoftTeamsMeeting']

    # 2026年3月改訂の業務記録の仕様に準じて追加(2026/3/23)
    # 本採用の業務記録簿の書式では一行めに「登録番号」と記載されている。これをCSVの旧書式に変更。
    # 停止したい場合は以下をコメントアウト
    ext_argv += ['--convert-old-style-tourokunum']

    flag = None

    try:
        argv = ext_argv + sys.argv[1:]

        argv, flag = libicsconvcsv.parse_args(argv, 1, short_opt, long_opt)
        if argv is None:
            __myhelp(exec_filename)

        Name = argv[0]
        if (len(Name)) == 0:
            raise ValueError("ERROR: 何らかの理由で引数Nameの取得に失敗しました。")

        # ファイル名に使えない記号検出
        # すでに同じような関数がありそうな気が若干するのだが(^_^;
        for i in list(Name):
            if i.isspace():
                raise ValueError("ERROR: 引数Nameに空白が含まれてます。")
            if not i.isprintable():
                raise ValueError("ERROR: 引数Nameにファイル名として使えない文字が含まれます。")
            if i in libicsconvcsv.ConstDat.FNAME_BAD_CHAR:
                raise ValueError(f"ERROR: 引数Nameにファイル名として使えない記号「{i}」が含まれます。")
    except ValueError as e:
        print("ERROR: ", e, file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit(1)

    #####################################################################
    # 出力ファイル名生成

    dt = datetime.datetime.now()

    if dt.month > 1:
        t = dt.year * 100 + (dt.month-1)
    else:
        t = (dt.year-1) * 100 + 12
    csv_fname_list[t] = f'./schedules{t:06}{Name}.csv'

    t = dt.year * 100 + dt.month
    csv_fname_list[t] = f'./schedules{t:06}{Name}.csv'

    #####################################################################
    # CSV変換
    # TODO: 単発で動かした場合と本プログラムで差分がないか確認
    # TODO: Windowsで確認。とくにファイルの日付確認。
    for key, value in csv_fname_list.items():
        libicsconvcsv.ics2csv(flag, __INPUT_ICS_FILENAME, value, key)

#End of main()
