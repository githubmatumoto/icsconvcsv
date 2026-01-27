# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025-2026 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
from enum import Enum
import sys
import io
import os
import re
import datetime
import zoneinfo
import csv
import getopt
import time
import dateutil
import vobject

HAIFU_URL = "https://qiita.com/qiitamatumoto/items/ab9e0cb9a6da257597a4"
GITHUB_URL = "https://github.com/githubmatumoto/icsconvcsv"

#######################################

HELP_LICENSE = f"""
=====================================================================
ライセンス:
Apache License 2.0

配布元:
{HAIFU_URL}
{GITHUB_URL}

修正履歴およびKnown bugsは CHANGELOG.md 参照ください。
"""
#######################################

__doc__ = """ICS to CSV コンバータ ライブラリ

ICS to CSV コンバータ icsconvcsv.py が利用するライブラリ。
"""
#######################################
__doc__ += HELP_LICENSE

#######################################
VERSION = "3.0"
#########################################################################

# utf_8_sig: WindowsでBOMをつける。
# BOMがあるかの確認は、od -a xx.csv | head とすると先頭に「o ; ?」が出力される。
CharSet = Enum('CharSet', [('utf_8', 'utf_8'),\
                           ('shift_jis', 'shift_jis'),\
                           ('utf_8_sig', 'utf_8_sig')])

# 終日スケジュールの日付の書式指定
AllDayFormat = Enum('AllDayFormat', \
                    [("addtime", "addtime"), ("today", "today"), ("nextday", "nextday"),\
                     ('todayremtime', 'todayremtime'), ('nextdayremtime', 'nextdayremtime')])

# CSVの出力指定
# 以下は内部向けなので、公開引数としない。
# omitdescriptionは outlookの全出力と、description無しのデータ(*-limit2)の比較用。
# cmpougaは garoonとoutlookのICS比較用
# debug1は原則変更しない。
CSVFormat = Enum('CSVFormat',\
                 [("simple", "simple"), ("garoon", "garoon"),\
                  ("outlookclassic", "outlookclassic"),\
                  ("cmpouga", "cmpouga"),\
                  ("omitdescription", "omitdescription"),\
                  ("debug1", "debug1"),\
                  ("debug2", "debug2")
                 ])

# 日時の出力形式。
# slash_ymd: 2025/12/31, 12:33  日本式順序。順番変更する場合は「F.output_sort」関連の修正必要。
# ISO_8601 basic: 20251231, 1233
# ISO_8601 extended: 2025/12/31, 12:33
# Ref: https://ja.wikipedia.org/wiki/ISO_8601
DateTimeFormat = Enum('DateTimeFormat', \
                      [("slash_ymd", "slash_ymd"), ("basic", "basic"), ("extended", "extended")])

class ConstDat:
    """各種定数を収納。"""
    # ヘッダ分割のデフォルト
    SPLIT_SUMMARY_HEAD = ('出張', '往訪', '来訪', '会議', '休み')
    # ヘッダ分割の拡張のデフォルト
    SPLIT_SUMMAEY_EXTEND_HEAD_GENERIC = "TODO,MEMO,授業,講義,実験,移動,TEST"

    BAD_CHAR = ['\\', '/', '[', ']', '<', '>', '?', '"', "'", "*", '-', '@', '{', '}']
    #ヘッダ分割時に問題になりそうな文字をヘッダとして拒否する
    # 「,」と「:」も使えないけど、セパレータなので除外。
    SPLIT_SUMMARY_HEAD_BAD_CHAR = BAD_CHAR
    # ファイル名で除外する記号。
    # カッコはLinuxでは特殊な意味をもつため。
    FNAME_BAD_CHAR = BAD_CHAR + ['(', ')']
    #
    # 特殊な文字列
    UNREF = "(REFERENCE DATA DOES NOT EXIST)"
    # 未定義を示す文字列。変更不可(これを書き換える場合は一部正規表現の修正が必要)
    NA = "(N/A)"

    #CSV必須ヘッダ
    CSV_REQUIRED = ["DTSTART:DAY", "DTSTART:TIME",\
                    "DTEND:DAY", "DTEND:TIME", "SUMMARY"]
    # 読み替え表
    CSV_TABLE_X_MICROSOFT_CDO_BUSYSTATUS = \
        {"WORKINGELSEWHERE":0, "TENTATIVE":1, "BUSY":2, "FREE":3, "OOF":4}

class FeatureFlags:
    """parse_argsなどで後で書き換える変数
    小文字は原則Bool型。大文字は原則Bool型以外"""
    def __init__(self):
        # 入力ファイルの日付確認、古いファイルへの警告を行うかいなか。
        # False:  日付確認を行わない。
        # True: 日付確認を行う
        self.old_file_check = False
        # 出力ファイルがすでに存在する場合、上書きするかの確認をするかしないか。
        # False: 上書き確認を行う
        # True: 上書き確認を行わない。
        self.overwrite = True

        #各要素の最後の改行と空白をすべて取り除く。
        # ics_parts_to_csv_buffer()
        self.remove_tail_cr = False
        # CSVに出力する時の各種処理関数
        #
        # CSVの1行めの最初のヘッダを、出力する(True)、しない(False)
        self.print_csv_header = False
        # 時刻の出力関係
        self.csv_show_timezone = False
        #メモ欄(description)の4行目以降を消す。
        self.description_delete_4th_line_onwards = False
        # Teamsの会議インフォメーションを消す。パスワードが入ってる。
        self.remove_teams_infomation = True
        # 登録番号(業務番号)の拡張フォーマットを使うか
        # True: 使う
        # False: 使わない
        self.enhanced_gyoumunum = False
        # VERSION1.3追加
        # 上書スケジュール(RECURRENCE-ID)の対応する。
        self.support_recurrence_id = True
        # 上書スケジュール(RECURRENCE-ID)で、
        # 基のスケジュールを隠す:True
        # 基のスケジュールを表示する: False. 基のスケジュールのSummaryに"Hidden: "と追加します。
        self.override_recurrence_id = True
        # CSVの出力の日付ソートを行う(True)。しない(False)
        self.output_sort = True
        # ICSのsummaryの分割を試みる。
        self.split_summary = True
        # EXDATEの書式問題で読み込みに失敗する問題を修正する。
        self.exdate_format_bugfix = True
        # EXDATEとUNTIでnativeとawareが混在するbugfixを有効にする。
        self.naive_aware_mixed_bugfix = True
        #
        # ZoneInfoで扱える文字列を指定する。
        # Ref: https://zenn.dev/fujimotoshinji/scraps/f9c25aeb00a716
        #
        # dtstart/dtendなどにTimeZoneが指定されていたらそれを優先。
        #
        self.OVERRIDE_TIMEZONE = None
        #self.OVERRIDE_TIMEZONE = "Asia/Tokyo" # 文字列
        #
        # 推測した値を保存。
        self.GUESS_TIMEZONE = None # 値は文字列ではない。TimeZoneオブジェクト
        self.guess_timezone_initalized = False
        #
        # 指定したUIDの細かい情報を表示する
        self.DEBUG_UID = None

        # CSVに出力する時の各種処理関数
        #
        self.CSV_ALLDAY_FORMAT = AllDayFormat.nextday
        self.CSV_DATE_TIME_FORMAT = None
        #
        # 予定の選択肢を追加
        # 追加項目。増やす場合はライブラリ呼び出し側から追加してください。
        self.SPLIT_SUMMARY_EXTEND_HEAD = []

        # CSVの出力フォーマット
        self.CSV_FORMAT = CSVFormat.simple

        # CSVのBODY(出力する部分)のCSVの各要素の位置
        self.CSV_POS = {}

        # CSVの内部処理用の(出力しない部分)を含めての各要素の位置.
        # 「H:LENGTH:」 CSVのHEADER(出力しない部分)の長さ
        # なので、次のようになる。
        # CSV_POS2["SUMMARY"] == (CSV_POS["SUMMARY"] + CSV_POS2["H:LENGTH"])
        #
        # 前述のCSVの1行めの最初のヘッダ(self.print_csv_header)と用語が混ざってますので注意ください。
        #
        self.CSV_POS2 = {"H:UID":0, "H:DTSTART":1, "H:RECURRENCE_ID":2, "H:LENGTH":3}

        self.CSV_HEADER = [ConstDat.NA, None, None] # 先頭部分のみ。 set_format()で後半をappendする。
        # == ["H:UID", "H:DTSTART", "H:RECURRENCE_ID"]

        # H:UIDはVEVENTのUID。CSVの項目一覧などは'(N/A)', RECURRENCE-IDで不可視化した場合はNoneを代入
        # Noneの場合はファイルへの出力対象外。
        # H:DTSTARTは datetime.datetime型もしくはdatetime.date型
        # H:RECURRENCE_IDは要素に含まれるならその値が入る。datetime.datetime型もしくはdatetime.date型
        # なければNone。datetime.datetimeの時はlocaltimeに変換する。

        # 試してないが、改行コードの話。
        # Ref: https://qiita.com/tatsuya-miyamoto/items/f57408064b803f55cf99

        # 試してないが、入力文字コードの変更はこちらが参考になる。
        # Ref: https://techblog.asahi-net.co.jp/entry/2021/10/04/162109
        #出力するCSVの文字コード
        self.CSV_ENCODING = None
        #

#######################################################
class TimeRange:
    """CSVの出力範囲を制限する処理をする関数"""
    @staticmethod
    def format_check(timerange: int) -> bool:
        """
        CSVの出力範囲を指定するtimerangeの値が異常な値でないかを判断する。

        引数:
        timerange (int): CSVの出力範囲を指定するtimerangeの値。

        返り値:
        正常ならTrue, 異常ならFalse
        """
        if timerange == 0:
            return True

        #timerange無効。全出力する。

        #異常値。
        if timerange < 0:
            return False

        #2000年未満/2100年以降は無効。
        y = timerange // 100
        if (y < 2000) or (y >= 2100):
            return False

        m = timerange % 100
        # 0月は無効/13月以上は無効。
        if (m < 1) or (m > 12):
            return False

        #異常なし。
        return True

    @staticmethod
    def guess_fname(FILENAME: str) -> int:
        """
        ファイル名をもとにCSVが出力する期間の推測を行います。

        引数:
        str: ファイル名

        返り値:
        推測したCSV出力期間。推測に失敗した場合はNoneを返します。

        """
        if re.search("all", FILENAME):
            return 0

        t = re.search(r"\d{6}", FILENAME)
        if not t:
            return None
        #print (f"DEBUG: {t}")
        ret = int(t.group())
        if not TimeRange.format_check(ret):
            return None
        return ret

    @staticmethod
    def guess(TIMERANGE: str, INPUT_ICS_FILENAME: str, OUTPUT_CSV_FILENAME: str) -> int:
        """
        CSVが出力する期間の値を推測します。

        引数:
        TIMERANGE:str
        INPUT_ICS_FILENAME:str
        OUTPUT_CSV_FILENAME:str

        返り値:
        CSV出力期間。失敗した場合はNoneを返します。
        """

        # CHECK TIMERANGE
        t_bak = TIMERANGE
        if TIMERANGE == "all":
            ret = 0
        elif TIMERANGE == "guessin":
            ret = TimeRange.guess_fname(INPUT_ICS_FILENAME)
            if ret is None:
                raise ValueError(f"ERROR: 入力ファイル名からCSVの期間の推測に失敗しました: {INPUT_ICS_FILENAME}")
        elif TIMERANGE == "guess":
            ret = TimeRange.guess_fname(OUTPUT_CSV_FILENAME)
            if ret is None:
                raise ValueError(f"ERROR: 出力ファイル名からCSVの期間の推測に失敗しました: {OUTPUT_CSV_FILENAME}")
        elif TIMERANGE.isdecimal():
            ret = int(TIMERANGE) + 0
        else:
            raise ValueError(f"ERROR: 期間指定の誤り: {t_bak}")

        if not TimeRange.format_check(ret):
            raise ValueError(f"ERROR: 期間指定の誤り: {t_bak}")
        return ret

    ###############################################################
    # timerange
    @staticmethod
    def is_collect(ics_time, timerange: int)->bool:
        """
        引数で渡した時刻ics_timeがCSVへの出力対象か判断します。

        引数:
        ics_time: datetime.datetime型もしくはdatetime.date型
        timerange: CSVの出力範囲を指定するtimerangeの値。

        返り値:
        出力対象ならTrue, それ以外はFalse

        """
        #print(f"DEBUG: timerange = {timerange}")
        #print(f"DEBUG: ics_time.year = {ics_time.year}")
        #print(f"DEBUG: ics_time.month = {ics_time.month}")
        if timerange == 0:
            return True

        if(ics_time.year == timerange//100) and (ics_time.month == timerange%100):
            return True
        return False
###
class Misc:
    """煩雑な関数"""
    @staticmethod
    def get_ics_val(ics_parts, name, default_val=None, exit_none=True):
        """
        vobjectのicsの要素を取り出す。

        引数:
        ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
        入ってる。

        name: 取り出す要素。例えば(summaryとかdtstartなど)
        default_val: nameで指定した要素が存在しない場合に返す値。
                 もしここにNoneを指定したときにnameで指定した
                 要素が存在しない場合は例外を送出する。

        exit_none  : 通常はdefault_valがNoneで要素が存在しない場合、
                 defaultでは停止するが、Falseだと停止せずに
                 Noneを返す。
    """
        name = name.lower()
        if hasattr(ics_parts, name):
            return getattr(ics_parts, name, default_val).valueRepr()

        if (default_val is None) and exit_none:
            print(f"ERROR: 不適切なICSファイルです。必須パラメータがありません: {name}", file=sys.stderr)
            ics_parts.prettyPrint()
            raise ValueError(f"ERROR: 不適切なICSファイルです。必須パラメータがありません: {name}")
        return default_val

    @staticmethod
    def csv_buffer_dump(buff: list, prefix="DEBUG:", uid=None, all_print=False, file=sys.stderr):
        """
        csv_buffeをdump

        uidで指定したcsvを出力します。
        uidが未指定の場合、all_print=Trueを指定するとすべてのcsvを出力します。
    """
        if F.DEBUG_UID is None:
            if not all_print:
                return

        print("----", file=file)
        for i in range(len(buff)):
            if uid is None:
                print(f"{prefix}{i}:{buff[i]}", file=file)
            else:
                if buff[i][F.CSV_POS2["H:UID"]] == uid:
                    print(f"{prefix}{i}:{buff[i]}", file=file)
        print("----", file=file)

class TZ:
    """ICSの時間関係やTimzeZoneの処理"""
    @staticmethod
    ###
    #########################################################################
    # 時間関係のis関数
    ###
    def is_aware(d) -> bool:
        """
        引数で渡した datetime オブジェクトd が aware(timezoneあり) かどうかを判定する
        """
        # Pythonマニュアルより:
        # date 型のオブジェクトは常に naive です。
        if type(d) is datetime.date:
            return False

        # Pythonマニュアルより:
        # 次の条件を両方とも満たす場合、 time オブジェクト t は aware です:
        # t.tzinfo が None でない
        # t.tzinfo.utcoffset(t) が None を返さない
        #どちらかを満たさない場合は、 t は naive です。

        if type(d) is datetime.datetime:
            if d.tzinfo is None:
                return False
            #MEMO: 旧版はd.tzinfo.utcoffset(None) と誤記してるので判断ミスする。
            if d.tzinfo.utcoffset(d) is None:
                return False
            return True

        raise RuntimeError(f"ERROR: 想定外の型が渡されました: type={type(d)}")

    ###
    @staticmethod
    def is_naive(d) -> bool:
        """
        引数で渡した datetime オブジェクトd が naive(timezoneなし) かどうかを判定する

        """
        return not TZ.is_aware(d)

    ###
    @staticmethod
    def is_am12(t) -> bool:
        """
        引数で渡した datetime オブジェクトt に 時/分/秒の時刻情報があり、
        深夜12時"00:00:00"ならTrue

        """
        if not type(t) is datetime.datetime:
            raise ValueError(f"ERROR: 想定外の型が渡されました。{type(t)}")
        return (t.hour + t.minute + t.second) == 0
    ###
    @staticmethod
    def hava_time(t) -> bool:
        """
        引数で渡した datetime オブジェクトt に 時刻情報があるかないか。

        MEMO: pylintに isinstance() を使うように指示されたが、
        なぜか動かなくなったので、もとに戻した
        """
        if type(t) is datetime.date:
            return False

        if type(t) is datetime.datetime:
            return True

        raise RuntimeError(f"ERROR: 想定外の型が渡された: = {type(t)}")

    @staticmethod
    def load_ics_vtimezone(ics_data: str):
        """
        TimeZoneデータ読み込み
        https://dateutil.readthedocs.io/en/stable/tz.html
        vobjectのサンプルtests.py

        返り値: None -> VTIMEZONEのデータが無かった。
        データ異常などの場合は、例外raiseで停止する。
    """
        ret = None

        try:
            ret = dateutil.tz.tzical(io.StringIO(ics_data))
        except ValueError as e:
            # Outlook(classic)が作ったICSを一気よみすると途中のVEVENTで
            # 解析に失敗する。
            new = []
            found_vtimezone = False
            found_vcalendar = False
            found_vevent = False
            for i in ics_data.splitlines():
                if re.match('BEGIN:VTIMEZONE', i):
                    found_vtimezone = True
                if re.match('BEGIN:VCALENDAR', i):
                    found_vcalendar = True
                if re.match('BEGIN:VEVENT', i):
                    if found_vevent:
                        raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(1)") from e
                    found_vevent = True
                    continue
                if re.match('END:VEVENT', i):
                    if not found_vevent:
                        raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(2)") from e
                    found_vevent = False
                    continue
                if found_vevent:
                    continue
                new.append(i)
            # end for i:

            if found_vevent:
                raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(3)") from e

            if not found_vcalendar:
                raise ValueError("ERROR: 入力して渡されたICSファイルにBEGIN:VCALENDARが無い") from e

            if not found_vtimezone:
                # VTIMEZONEが無い。おそらくFloating timeモード。ただし時刻がUTC(世界標準時)の可能性あり。
                ret = None
            else:
                #print("\n".join(new))
                ret = dateutil.tz.tzical(io.StringIO("\n".join(new) + "\n"))

        #print(type(ret))
        #if len(ret.keys()) == 1:
            #t = datetime.datetime(2003, 9, 27, 12, 40, 12, tzinfo=ret.get())
            #t = datetime.datetime(2003, 9, 27, 12, 40, 12, tzinfo=dateutil.tz.tzutc())
            #print(t, file=sys.stderr)

        #print(tzs.get('Tokyo Standard Time'))
        #jp = vobject.icalendar.TimezoneComponent(tzs.get('Tokyo Standard Time'))
        #print(jp)

        return ret

    ###
    #
    #
    @staticmethod
    def guess_timezone_init(cal_tz: dateutil.tz.tz.tzical, override_timezone: str = None):
        """
        TimeZoneを推測する関数の初期化

        制御変数:
        G.OVERRIDE_TIMEZONE:ただし ics2csv()で使ってます。

        """
        global F

        F.guess_timezone_initalized = False

        n = []
        if not n is None:
            n = cal_tz.keys()

        if not override_timezone is None:
            print("INFO: 引数でデフォルトのTimeZoneが指定されています。", file=sys.stderr)

                # VTIMEZONEで定義されていたTimeZoneから探す。
            if override_timezone in n:
                F.GUESS_TIMEZONE = cal_tz.get(override_timezone)
            else:
                # OS定義のTimeZOneから探す。
                try:
                    F.GUESS_TIMEZONE = zoneinfo.ZoneInfo(override_timezone)
                except zoneinfo.ZoneInfoNotFoundError as e:
                    raise ValueError(f"ERROR: 無効なTimeZoneとして[{override_timezone}]が指定されました。") from e

            print(f"INFO: TimeZoneとして[{override_timezone}]を採用します。", file=sys.stderr)
            F.guess_timezone_initalized = True
            return

        if len(n) == 0:
            print("INFO: ICSデータにTimeZoneデータがありません。", file=sys.stderr)
            print("INFO: Floating Timeのデータです。(Ref: RFC5545, 3.3.12. TIME)", file=sys.stderr)
            F.GUESS_TIMEZONE = None
            F.guess_timezone_initalized = True
            return

        if len(n) == 1:
            F.GUESS_TIMEZONE = cal_tz.get()
            F.guess_timezone_initalized = True
            return

        if len(n) > 1:
            print("INFO: ICSファイルにTimzeZoneが複数定義されています。", file=sys.stderr)
            print(f"INFO: 現在定義されているTimeZone一覧: {cal_tz.keys()}", file=sys.stderr)
            print(f"INFO: TimeZoneとして1番目に定義されている[{n[0]}]を採用します。", file=sys.stderr)
            print("WARNING: 採用したTimeZoneが不適切な場合、日時計算に失敗します。", file=sys.stderr)
            print("WARNING: 誤ったCSVが生成される場合は引数-TでTimeZoneを指定してください。", file=sys.stderr)
            F.GUESS_TIMEZONE = cal_tz.get(n[0])
            F.guess_timezone_initalized = True
            return

        raise ValueError("ERROR: ここには来ないはずだが。。")

    @staticmethod
    def guess(exit_error: bool = True):
        """
        TimeZoneを推測する。事前にinit_guess_timezone()で初期化する必要あり。

        exit_error:
               False:TimeZoneの推測に失敗してもエラーとしない。
               True:TimeZoneの推測に失敗したら終了。

        """
        if not F.guess_timezone_initalized:
            raise ValueError("ERROR: 初期化されていません")

        if F.GUESS_TIMEZONE is None:
            if exit_error:
                raise ValueError("ERROR: 壊れたICSファイルです。恐らくFloating Timeだが世界標準時が使われているためローカルタイムに変換できない。")
            return None

        return F.GUESS_TIMEZONE

    #
    @staticmethod
    def load_ics(ics_data: str, override_timezone: str = None):
        """
        STRのICSデータからTimeZone関連の情報を読みこみTimeZone関係の初期化
        """
        cal_tz = TZ.load_ics_vtimezone(ics_data)
        TZ.guess_timezone_init(cal_tz, override_timezone)

    #########################################################################
    # TimeZoneの変換関係の関数。
    ###
    @staticmethod
    def naive2aware(d, exit_error: bool = True) -> datetime.datetime:
        """
        RFC5545ではfloating timeという概念があるdatetimeのnaive timeとほぼ同等。

        引数で渡したdatetimeオブジェクトdはnaiveな
        (datetime.datetimedate or datetime.date)とする。

        引数dをawareなdatetimeに変換する。

        時刻情報が無い日付のみであるdatetime.date型を渡された時は0時0分0秒とする。

        exit_error:
               False:TimeZoneの推測に失敗してもエラーとしない。
               True:TimeZoneの推測に失敗したら終了。

        """
        #print(f"DEBUG(conver_aware/pre):{d}", file=sys.stderr)
        #print(f"DEBUG type(d) = {type(d)}", file=sys.stderr)

        if TZ.is_aware(d):
            #print(f"DEBUG(conver_aware/aft):{d}", file=sys.stderr)
            return d

        y = d.year
        m = d.month
        day = d.day

        s = minute = h = 0
        if type(d) is datetime.datetime:
            s = d.second
            minute = d.minute
            h = d.hour
            #print(f"DEBUG s={s}, min={minute}, h={h}", file=sys.stderr)

        tz = TZ.guess(exit_error)
        #tzinfoのdefault引数はNoneだからtzの値がNoneであっても調べてない。
        d = datetime.datetime(y, m, day, h, minute, s, tzinfo=tz)
        #print(f"DEBUG(conver_aware/aft):{d}", file=sys.stderr)
        return d

    @staticmethod
    def to_localtime(d, exit_none=True, exit_naive=False):
        """

        引数で渡したdatetimeオブジェクトdに TimeZone情報があれば、ローカルタイ
        ムに変換する。

        TimeZone情報がなければ、何もしない。

        exit_none: Noneが渡された時の挙動。
               True: 例外を送出。
               False : Noneを返す。

        exit_naive: Timezoneがないnaive timeを渡された時の挙動。
               True: 例外を送出。
               False : なにもせずに渡された時刻を返す。

        """
        if d is None:
            if exit_none:
                raise RuntimeError("ERROR: 引数にNoneが渡されました")
            return None

        if TZ.is_aware(d):
            return d.astimezone(TZ.guess()) # ローカルタイムに変換。

        if exit_naive:
            raise RuntimeError(f"ERROR: floatingtimeをローカルタイムへの変換しようとしました, time={d}")
        return d


    ##########################################################################
    ###
    @staticmethod
    def ics_parts_to_csv_time(ics_parts, rrule_start) -> tuple:
        """VEVENTのDTSTART&DTENDをCSV出力用の文字列(tuple)に変換する。

       引数:

        ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
                   入ってる。

        rrule_start: 繰返しスケジュールの時の開始時刻。VEVENTのDTSTART
                     とDTENDを置き換える。Noneの場合は、VEVENTの
                     DTSTARTとDTENDがそのまま使われる。
                     datetime.datetime型もしくはdatetime.date型

       返り値:
        タプルで5つ
          ("開始日","開始時刻","終了日","終了時刻", "終日スケジュールフラグ")

        終日スケジュールフラグ:
              フラグ’--allday-format-add-time’指定時のみ意味がある。
              時間情報がある(0時開始/翌日0時終了)終日スケジュールの場合はFalse
              時間情報がない終日スケジュールの場合はTrue

       外部制御変数:
        F.show_timezone: Bool型
             True: TimeZone情報が有る場合、TimeZone情報込で出力。
                TimeZone情報あり: 12:34:55+09:00, 12:34:55-03:00
                TimeZone情報なし: 12:34:55, 12:34:55

             False: TimeZone情報がある場合、関数TZ.guess()を使って
                推測したTimeZoneに修正の上、TimeZone情報を削って表示
                12:34:55, 12:34:55

        F.CSV_ALLDAY_FORMAT: enum AllDayFormat型
             ヘルプの　引数 --allday-format-XXX の欄参照。

        """
        global F

        start = Misc.get_ics_val(ics_parts, 'x-org-dtstart', None, exit_none=False)
        if start is None:
            start = Misc.get_ics_val(ics_parts, 'dtstart')

        end = Misc.get_ics_val(ics_parts, 'x-org-dtend', None, exit_none=False)
        if end is None:
            end = Misc.get_ics_val(ics_parts, 'dtend')

        if False:
            print(f"DEBUG: rrule_start = {rrule_start}", file=sys.stderr)
            print(f"DEBUG: start = {start}", file=sys.stderr)
            print(f"DEBUG: end = {end}", file=sys.stderr)
            print(f"DEBUG: F.csv_show_timezone = {F.csv_show_timezone}", file=sys.stderr)
            print(f"DEBUG: F.CSV_ALLDAY_FORMAT = {F.CSV_ALLDAY_FORMAT}", file=sys.stderr)
            print(f"DEBUG: F.CSV_FORMAT = {F.CSV_FORMAT}", file=sys.stderr)

        if not rrule_start is None:
            if TZ.is_aware(start) != TZ.is_aware(rrule_start):
                #この状態になる場合は、事前処理にミスしてる可能性大
                raise ValueError("BUG: timezoneありなし混在")
            end = rrule_start+(end-start)
            start = rrule_start

        # 時刻の出力形式
        date_f = "%Y/%m/%d" # DateTimeFormat.slash_ymd
        time_f = "%H:%M:%S" # DateTimeFormat.slash_ymd

        if F.CSV_DATE_TIME_FORMAT == DateTimeFormat.basic:
            date_f = "%Y%m%d"
            time_f = "%H%M%S"

        if F.CSV_DATE_TIME_FORMAT == DateTimeFormat.extended:
            date_f = "%Y-%m-%d"

        if False:
            print(f"DEBUG: start = {start}", file=sys.stderr)
            print(f"DEBUG: end = {end}", file=sys.stderr)

        # ローカルタイムに変換。時間情報がnaiveの時は何もしない。
        start = TZ.to_localtime(start)
        end = TZ.to_localtime(end)

        if False:
            print(f"DEBUG: start(localtime) = {start}", file=sys.stderr)
            print(f"DEBUG: end(localtime) = {end}", file=sys.stderr)

        all_day = None
        if TZ.hava_time(start): # 時刻情報あり
            all_day = False
            if F.CSV_ALLDAY_FORMAT in (AllDayFormat.nextdayremtime, AllDayFormat.todayremtime):
                # ローカルタイムで0:00の時に時間を除去する。
                if TZ.is_am12(start) and  TZ.is_am12(end):
                    start = start.date()
                    end = end.date()
        else: # 時刻情報なし
            all_day = True
            if F.CSV_ALLDAY_FORMAT == AllDayFormat.addtime:
                #print(f"DEBUG: IN addtime", file=sys.stderr)

                if start == end:
                    raise ValueError("ICSデータ異常startとendに時刻がなく、start == end")

                #第2引数のFalseはTZがないデータでもエラーとしない指示。
                start = TZ.naive2aware(start, False)
                end = TZ.naive2aware(end, False)
                #print(f"DEBUG: conv start = {start}", file=sys.stderr)
                #print(f"DEBUG: conv end = {end}", file=sys.stderr)

        if False:
            print(f"DEBUG: start = {start}", file=sys.stderr)
            print(f"DEBUG: end = {end}", file=sys.stderr)


        if TZ.hava_time(start):
            if F.csv_show_timezone:
                if TZ.is_aware(start):
                    time_f += "%z"
            return start.strftime(date_f), \
                 start.strftime(time_f), \
                 end.strftime(date_f), \
                 end.strftime(time_f), \
                 all_day

        #以下は時刻情報がない場合の処理

        if F.CSV_ALLDAY_FORMAT in (AllDayFormat.today, AllDayFormat.todayremtime):
            end = end - datetime.timedelta(days=1)

        return start.strftime(date_f), "", end.strftime(date_f), "", all_day
    ###

class PreSetup:
    """引数の処理およびICSの前処理関係の関数"""
    @staticmethod
    ###
    #########################################################################
    # vobjectがRRULEのEXDATE関連で例外を履く記述の修正関数
    #
    def set_format(override_encoding: CharSet, override_all_day_format: AllDayFormat, \
                   override_datetime_format: DateTimeFormat):
        """
        対応CSVフォーマットの初期設定および出力文字コードの設定。
    """
        # 各CSVフォーマットの初期値を設定する。
        # garoonの時のみshift_jis
        if F.CSV_FORMAT == CSVFormat.garoon:
            F.CSV_ENCODING = CharSet.shift_jis
            F.CSV_ALLDAY_FORMAT = AllDayFormat.today
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.slash_ymd
        elif F.CSV_FORMAT == CSVFormat.simple:
            F.CSV_ENCODING = CharSet.utf_8
            F.CSV_ALLDAY_FORMAT = AllDayFormat.nextday
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.extended
        elif F.CSV_FORMAT == CSVFormat.outlookclassic:
            F.CSV_ENCODING = CharSet.utf_8
            F.CSV_ALLDAY_FORMAT = AllDayFormat.addtime
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.slash_ymd
            override_all_day_format = None # 上書き不可。
        elif F.CSV_FORMAT == CSVFormat.cmpouga:
            # Outlookと GaroonのICS比較用。
            F.CSV_ENCODING = CharSet.utf_8
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.slash_ymd
            F.CSV_ALLDAY_FORMAT = AllDayFormat.todayremtime
            F.remove_tail_cr = True
            F.enhanced_gyoumunum = True
            ModCSV.set_summary_extend_head("おそらくバグ:")
        elif F.CSV_FORMAT == CSVFormat.omitdescription:
            F.CSV_ENCODING = CharSet.utf_8
            #GaroonとOutlook比較時はこれを入れないと古いデータの
            F.CSV_ALLDAY_FORMAT = AllDayFormat.todayremtime
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.extended
        elif F.CSV_FORMAT == CSVFormat.debug1:
            # debug1は原則変更しない。debug2以降は頻繁に変更の可能性あり。
            F.CSV_ENCODING = CharSet.utf_8
            F.CSV_ALLDAY_FORMAT = AllDayFormat.today
            F.CSV_DATE_TIME_FORMAT = DateTimeFormat.extended

        # 文字コードを上書きする。
        if not override_encoding is None:
            F.CSV_ENCODING = override_encoding

        # 終日スケジュールの日付の書式指定を上書きする。
        if not override_all_day_format is None:
            F.CSV_ALLDAY_FORMAT = override_all_day_format

        if not override_datetime_format is None:
            F.CSV_DATE_TIME_FORMAT = override_datetime_format

        #CSVの項目の位置などの初期化。
        # 「：」が入ってるのは特殊な値。
        # CSVのlistの実際の位置はF.CSV_B_OFFSET+F.CSV_POS[HOGEHOGE]になります。

        # 独自定義のICS要素の追加手順:

        # ICSファイルの要素ABCの中身を加工せずにそのまま出力する場合は
        #  F.CSV_POS["ABC"]  = CSVの位置
        # と記載します。何らかの加工をする場合は「:」付きで以下のような感じで記載し
        #  F.CSV_POS["X:ABC"]  = CSVの位置
        #  F.CSV_POS["ABC:適当な名前"]  = CSVの位置
        # 関数Main.ics_parts_to_csv_buffer()にその要素の
        # 処理方法を記載します。

        if F.CSV_FORMAT in (CSVFormat.garoon, CSVFormat.cmpouga):
            h_tail = ["開始日", "開始時刻", "終了日", "終了時刻", "予定", "予定詳細", "メモ"]
            # 「：」が入ってるのは特殊な値。
            F.CSV_POS["DTSTART:DAY"] = 0  # DTSTARTの日付
            F.CSV_POS["DTSTART:TIME"] = 1 # DTSTARTの時刻
            F.CSV_POS["DTEND:DAY"] = 2
            F.CSV_POS["DTEND:TIME"] = 3
            F.CSV_POS["SUMMARY:H"] = 4    # SUMMARYのヘッダ分離
            F.CSV_POS["SUMMARY"] = 5
            F.CSV_POS["DESCRIPTION"] = 6
            F.CSV_POS2["B:LENGTH"] = 7 # CSVの項目の長さ
        elif F.CSV_FORMAT == CSVFormat.omitdescription:
            h_tail = ["DTSTART:DAY", "DTSTART:TIME", "DTEND:DAY", "DTEND:TIME", "SUMMARY"]
            F.CSV_POS["DTSTART:DAY"] = 0  # DTSTARTの日付
            F.CSV_POS["DTSTART:TIME"] = 1 # DTSTARTの時刻
            F.CSV_POS["DTEND:DAY"] = 2
            F.CSV_POS["DTEND:TIME"] = 3
            F.CSV_POS["SUMMARY"] = 4
            F.CSV_POS2["B:LENGTH"] = 5 # CSVの項目の長さ
        elif F.CSV_FORMAT == CSVFormat.debug1:
            # debug1は原則変更しない。debug2以降は頻繁に変更の可能性あり。
            h_tail = ["開始日", "開始時刻", "終了日", "終了時刻", "終日イベント", "予定", "予定詳細"]
            F.CSV_POS["DTSTART:DAY"] = 0  # DTSTARTの日付
            F.CSV_POS["DTSTART:TIME"] = 1 # DTSTARTの時刻
            F.CSV_POS["DTEND:DAY"] = 2
            F.CSV_POS["DTEND:TIME"] = 3
            F.CSV_POS["X:ALLDAY_EVENT"] = 4
            F.CSV_POS["SUMMARY:H"] = 5    # SUMMARYのヘッダ分離
            F.CSV_POS["SUMMARY"] = 6
            F.CSV_POS2["B:LENGTH"] = 7 # CSVの項目の長さ
        elif F.CSV_FORMAT == CSVFormat.simple:
            h_tail = ["DTSTART:DAY", "DTSTART:TIME", "DTEND:DAY", "DTEND:TIME", "SUMMARY",\
                      "DESCRIPTION", "X-MICROSOFT-CDO-BUSYSTATUS", "CATEGORIES"]
            F.CSV_POS["DTSTART:DAY"] = 0
            F.CSV_POS["DTSTART:TIME"] = 1
            F.CSV_POS["DTEND:DAY"] = 2
            F.CSV_POS["DTEND:TIME"] = 3
            F.CSV_POS["SUMMARY"] = 4
            F.CSV_POS["DESCRIPTION"] = 5
            F.CSV_POS["X-MICROSOFT-CDO-BUSYSTATUS"] = 6
            F.CSV_POS["CATEGORIES"] = 7
            F.CSV_POS2["B:LENGTH"] = 8
        elif F.CSV_FORMAT == CSVFormat.outlookclassic:
            h_tail = ["件名", "開始日", "開始時刻", "終了日", "終了時刻", "終日イベント",\
                      "アラーム オン/オフ", "アラーム日付", "アラーム時刻", "会議の開催者",\
                      "必須出席者", "任意出席者", "リソース", "プライベート", "経費情報",\
                      "公開する時間帯の種類", "支払い条件", "場所", "内容", "秘密度", \
                      "分類", "優先度"]
            F.CSV_POS["SUMMARY"] = 0
            F.CSV_POS["DTSTART:DAY"] = 1
            F.CSV_POS["DTSTART:TIME"] = 2
            F.CSV_POS["DTEND:DAY"] = 3
            F.CSV_POS["DTEND:TIME"] = 4
            F.CSV_POS["X:ALLDAY_EVENT"] = 5
            F.CSV_POS["ORGANIZER:CN"] = 9
            F.CSV_POS["ATTENDEE:CN:RSVP:TRUE"] = 10
            F.CSV_POS["ATTENDEE:CN:RSVP:FALSE"] = 11
            F.CSV_POS["X-MICROSOFT-CDO-BUSYSTATUS:NUM"] = 15
            F.CSV_POS["DESCRIPTION"] = 18
            F.CSV_POS2["B:LENGTH"] = 22
        else:
            raise ValueError("Internal Error: テーブルの初期化失敗(1)")

        #各種チェック
        #if len(F.CSV_POS) != F.CSV_POS2["B:LENGTH"]:
        #raise ValueError("Internal Error: テーブルの初期化失敗(2)")

        # ヘッダの生成
        if len(F.CSV_HEADER) != F.CSV_POS2["H:LENGTH"]:
            raise ValueError("Internal Error: テーブルの初期化失敗(3)")


        F.CSV_HEADER += h_tail

        for k, v in F.CSV_POS.items():
            F.CSV_POS2[k] = v + F.CSV_POS2["H:LENGTH"]

        #必須確認
        for i in ConstDat.CSV_REQUIRED:
            if i not in F.CSV_POS:
                raise ValueError(f"Internal Error: テーブルの初期化失敗(必須項目{i}なし)")

        # 共存不可の構成になってないか確認
        if F.enhanced_gyoumunum and ('DESCRIPTION' not in F.CSV_POS):
            raise ValueError("Internal Error: テーブルの初期化失敗(必須項目なし DESCRIPTION)")

        if 'SUMMARY:H' not in F.CSV_POS:
            F.split_summary = False

        if 'DESCRIPTION' not in F.CSV_POS:
            F.description_delete_4th_line_onwards = False
            F.remove_teams_infomation = False
            F.enhanced_gyoumunum = False

        #if 'X:ALLDAY_EVENT' in F.CSV_POS:
        #    if F.CSV_ALLDAY_FORMAT != AllDayFormat.addtime:
        #        raise ValueError("Internal Error: テーブルの初期化失敗(未対応の組み合わせ)")
    #####
    @staticmethod
    def parse_args(argv: list, amari_argv: int = -1, \
                   allow_short_opt: str = None, allow_long_opt: list = None) -> list:
        """
        オプションの解析を行います。
        多数オプションがありますので、制限されたコマンドのために
        allow_short_optおよびallow_long_optを使います。与えられた
        引数のみ有効にします。

        amari_argvはマイナスがついた引数解析後に残るマイナスがない引数の数です。

        オプション解析に失敗するとNoneを返します。

        返り値は、引数解析後に残るマイナスがない引数です。

        """
        global F

        # 引数の制限をする仕様をいれてるので、 -X=YY 形式の引数は --X-AA, --X-BBなど
        # で直接指定可能にしている。

        short_opt = "h"
        long_opt = ["help"]
        #
        short_opt += "C:"
        long_opt += ["char-set="]
        #
        short_opt += "F:"
        long_opt += ["format-simple", "format-garoon"]
        long_opt += ["format-classic", "format-outlookclassic"]
        #
        short_opt += "T:"
        long_opt += ["override-timezone="]
        #
        short_opt += "s"
        long_opt += ["disable-split-summary"]
        #
        short_opt += "m"
        long_opt += ["extend-summary-head"]
        #
        long_opt += ["add-summary-head="]
        #
        short_opt += "z"
        long_opt += ["enhance-gyoumunum", "enhance-gyoumu-number"]
        long_opt += ["enhance-tourokunum", "enhance-touroku-number"]
        #
        short_opt += "k"
        long_opt += ["print-csv-header"]
        #
        long_opt += ["show-timezone"]
        #最後に指定されたオプションが有効
        long_opt += ["allday-format-today", "allday-format-nextday"]
        long_opt += ["allday-format-am12", "allday-format-addtime"]
        long_opt += ["allday-format-today-remove-time", "allday-format-nextday-remove-time"]

        #最後に指定されたオプションが有効
        long_opt += ["day-format-iso8601-basic", "day-format-iso8601-extended",\
                     "day-format-slash-ymd",]
        #
        long_opt += ["delete-4th-line-onward", "show-teams-infomation"]
        long_opt += ["remove-tail-cr", "show-hidden-schedules", "disable-recurrence-id"]
        long_opt += ["disable-exdate-format-bugfix", "disable-naive-aware-mixed-bugfix"]
        long_opt += ["DEBUG-UID="]
        #
        #最後に指定されたオプションが有効
        short_opt += "W"
        long_opt += ["disable-file-exist-test", "enable-file-exist-test"]

        # 有効な引数の上書き。
        if not allow_long_opt is None:
            #print(allow_long_opt)
            long_opt = allow_long_opt

        if not allow_short_opt is None:
            #print(allow_short_opt)
            short_opt = allow_short_opt

        opts, argv = getopt.gnu_getopt(argv, short_opt, long_opt)

        override_encoding = None
        override_all_day_format = None
        override_datetime_format = None
        for o, a in opts:
            if o in ("-C", "--char-set"):
                override_encoding = re.sub(r'\-', '_', a.lower())
                for e in CharSet:
                    if e.name == override_encoding:
                        override_encoding = e
                if not override_encoding in CharSet:
                    raise ValueError(f"ERROR: 未対応の文字コードです: {a}")
            elif o == "-F":
                F.CSV_FORMAT = a
                #ハイフンを取り除く
                F.CSV_FORMAT = re.sub(r'\-', '', F.CSV_FORMAT.lower())

                for e in CSVFormat:
                    if e.name == F.CSV_FORMAT:
                        F.CSV_FORMAT = e
                if not type(F.CSV_FORMAT) is CSVFormat:
                    raise ValueError(f"ERROR: 未対応のCSV Formatです: {F.CSV_FORMAT}")
            elif o == "--format-simple":
                F.CSV_FORMAT = CSVFormat.simple
            elif o == "--format-garoon":
                F.CSV_FORMAT = CSVFormat.garoon
            elif o in ("--format-outlook-classic", "--format-outlookclassic"): # 未実装
                F.CSV_FORMAT = CSVFormat.outlookclassic # ハイフンなし
            elif o in ("-T", "--override-timezone"):
                F.OVERRIDE_TIMEZONE = a
            elif o in ("-s", "--disable-split-summary"):
                F.split_summary = False
            elif o in ("-m", "--extend-summary-head"):
                ModCSV.set_summary_extend_head("おそらくバグ:")
            elif o == "--add-summary-head":
                ModCSV.set_summary_extend_head("引数--add-summary-headに", a)
            elif o in ("-z", "--enhance-gyoumunum", "--enhance-gyoumu-number",\
                       "--enhance-tourokunum", "--enhance-touroku-number"):
                F.enhanced_gyoumunum = True
            elif o in ("-k", "--print-csv-header"):
                F.print_csv_header = True
            elif o == "--show-timezone": # old opt: -t
                F.csv_show_timezone = True
            elif o in ("--allday-format-addtime", \
                       "--allday-format-am12"): # old opt: -o
                override_all_day_format = AllDayFormat.addtime
            elif o == "--allday-format-today":
                override_all_day_format = AllDayFormat.today
            elif o == "--allday-format-nextday":
                override_all_day_format = AllDayFormat.nextday
            elif o == "--allday-format-today-remove-time": # old opt: -g
                override_all_day_format = AllDayFormat.todayremtime
            elif o == "--allday-format-nextday-remove-time":
                override_all_day_format = AllDayFormat.nextdayremtime
            elif o == "--day-format-slash-ymd":
                override_datetime_format = DateTimeFormat.slash_ymd
            elif o == "--day-format-iso8601-basic":
                override_datetime_format = DateTimeFormat.basic
            elif o == "--day-format-iso8601-extended":
                override_datetime_format = DateTimeFormat.extended
            elif o == "--delete-4th-line-onward": # old opt: -d
                F.description_delete_4th_line_onwards = True
            elif o == "--show-teams-infomation": # old opt: -p
                F.remove_teams_infomation = False
            elif o == "--remove-tail-cr": # old opt -r
                F.remove_tail_cr = True
            elif o == "--show-hidden-schedules": # old opt -w
                F.override_recurrence_id = False
            elif o == "--disable-recurrence-id": # old opt -x
                F.support_recurrence_id = False
            elif o == "--disable-exdate-format-bugfix": # old opt: -b
                F.exdate_format_bugfix = False
            elif o == "--disable-naive-aware-mixed-bugfix":
                F.naive_aware_mixed_bugfix = False
            elif o == "--DEBUG-UID":
                F.DEBUG_UID = a
            elif o == "--enable-file-exist-test":
                # 引数の指定順序依存あり。
                # 出力ファイルの上書き確認/入力ファイルの日付確認を行なう。
                F.overwrite = False
                F.old_file_check = True
            elif o in ("-W", "--disable-file-exist-test"):
                # 引数の指定順序依存あり。
                # 出力ファイルの上書き確認/入力ファイルの日付確認を行わない
                F.overwrite = True
                F.old_file_check = False
            elif o in ("-h", "--help"):
                return None

        if (amari_argv != -1) and (len(argv)) != amari_argv:
            raise ValueError("ERROR: 引数を間違えてます。")

        PreSetup.set_format(override_encoding, override_all_day_format, override_datetime_format)
        return argv

    ###
    @staticmethod
    def find_ics_data(data: list, key: str, exit_none=True) -> int:
        """
        文字列型で渡されたVEVENTのデータdataからkeyで指示された
        要素がある行を探します。要素が見つからない場合exit_none=True
        の場合は例外を送出します。
        """
        for i, d in enumerate(data):
            if re.match(key, d):
                return i
        if exit_none:
            raise RuntimeError(f"ICSのアイテム「{key}」が無いVEVENTが渡された")
        return -1

    ###
    @staticmethod
    def bugfix_exdate_format_aux(data: list) -> list:
        """
        bugfix_exdate_formatの補助関数1

        BEGIN:VEVENTからEND:VEVENTの間のデータをSTRING型のlistで渡して、
        EXDATE関連を修正する。

        """
        flag_exdate = PreSetup.find_ics_data(data, 'EXDATE')

        # EXDATEの時刻指定は複数ある場合があるので、修正時は要注意。
        # 「EXDATE:20250909,20250915」など。

        exdate = data[flag_exdate].split(':')
        if len(exdate) != 2:
            raise ValueError(f"ERROR: DXDATEの書式異常: {data[flag_exdate]}")

        # 時刻情報(T)が一つでもあれば何もしない。
        #「EXDATE:20250909T112233」など。
        if re.search('T', exdate[1]):
            return data

        # EXDATEに何らかのオプションがあれば何もしない。
        if not re.fullmatch('EXDATE', exdate[0]):
            return data

        # maybe Garoon. 時刻情報で(時・分)なし。
        # search: 「EXDATE:20250909」で後ろにTなし。
        exdate[0], count = re.subn('^EXDATE$', 'EXDATE;VALUE=DATE', exdate[0])
        if count == 0:
            raise RuntimeError(f"ERROR: 多分バグ(Garoonfix): {exdate[0]}")
        data[flag_exdate] = exdate[0] + ":" + exdate[1]
        return data

    ###
    @staticmethod
    def bugfix_exdate_format(data: str) -> str:
        """EXDATE関連のbugfix。ICSのファイルをすべて読み込んだ
    string型のdataを渡して、修正して返却する。

    RRULEのEXDATEが下記形式だとライブラリvobject-0.99では例外を
    送出します。

      EXDATE:20251128

    本関数は、下記形式に修正します。

      EXDATE;VALUE=DATE:20251128

      バグの詳細についてはmisc/TECH-MEMO.txt 参照ください。


        """
        flag_debug = False
        #flag_debug = True
        #print(f"DEBUG: arg_type:  {type(data)}")
        # typeはstrを想定
        if not type(data) is str:
            raise RuntimeError(f"ERROR: 想定外の型が渡されました: type={type(data)}")

        # https://maku77.github.io/python/numstr/split-lines.html
        # 文字列を改行で分割する。
        lines = data.splitlines()
        org_line_num = len(lines)

        flag_in_vevent = False
        flag_hava_exdate = False
        lines_in_vevent = []

        lines_ret = []

        for i in lines:
            if re.match('BEGIN:VEVENT', i):
                if flag_in_vevent is True:
                    raise RuntimeError("ERROR: 「BEGIN:VEVENT」が二重に現れました")
                if len(lines_in_vevent) != 0:
                    raise RuntimeError("ERROR: おそらくバグ")
                flag_in_vevent = True
                lines_in_vevent.append(i)
                continue

            if re.match('END:VEVENT', i):
                if flag_in_vevent is False:
                    raise RuntimeError("ERROR: 「END:VEVENT」が二重に現れました")
                flag_in_vevent = False
                lines_in_vevent.append(i)
                if flag_hava_exdate:
                    if flag_debug:
                        print("DEBUG: PRE:--EXDATE--\n"+'\n'.join(lines_in_vevent) \
                              + "-----\n", file=sys.stderr)
                    lines_in_vevent = PreSetup.bugfix_exdate_format_aux(lines_in_vevent)
                    if flag_debug:
                        print("DEBUG: AFT:----\n"+'\n'.join(lines_in_vevent) \
                              + "-----\n", file=sys.stderr)
                    flag_hava_exdate = False

                lines_ret += lines_in_vevent
                lines_in_vevent = []
                continue

            if not flag_in_vevent:
                lines_ret.append(i)
                continue

            if re.match('EXDATE', i):
                flag_hava_exdate = True

            lines_in_vevent.append(i)
        ##
        if len(lines_ret) != org_line_num:
            raise RuntimeError("ERROR: 行数が変化している。たぶんバグ")

        return "\n".join(lines_ret) + "\n"

class FileIO:
    """ファイルの読み書き関連"""
    @staticmethod
    def file2str(fname: str) -> str:
        """
        ファイルから読み込み、文字列型で返します。
    """
        ret = ""
        if fname == "stdout" or fname[0] == "-":
            #raise RuntimeError(f"入力ファイル")
            print(f"ERROR: 入力元のICSファイル名指定エラー: {fname}", file=sys.stderr)
            sys.exit(1)

        if fname == "stdin":
            ret = sys.stdin.read()
            return ret

        if not os.path.exists(fname):
            print(f"ERROR: 入力元のICSファイル「{fname}」が存在しません", file=sys.stderr)
            sys.exit(1)

        global F
        if F.old_file_check:
            stat_info = os.stat(fname)
            current_time = time.time()
            diff_time = current_time - stat_info.st_mtime

            #print("mtime=", stat_info.st_mtime)
            #print("curre=", current_time)
            #print("diff =", diff_time)

            # 日付が未来。プログラム停止。
            if diff_time <= -60:
                print(f"ERROR: 入力元のICSファイル「{fname}」の日付が未来です。", file=sys.stderr)
                print("ERROR: パソコンの時計が正しくない可能性があります。", file=sys.stderr)
                sys.exit(1)

                # 日付が3時間以上前。プログラム停止。
            if diff_time >= 60*60*3:
                h = int(diff_time/3600 + 0.5)
                print(f"ERROR: 入力元のICSファイル「{fname}」の日付が古いです。", file=sys.stderr)
                print(f"ERROR: {h}時間前のファイルです。最新のファイルをダウンロードしてください。", file=sys.stderr)
                print("ERROR: 最新の場合はファイル名が間違えてないか確認ください。", file=sys.stderr)
                sys.exit(1)

            old_m = 15
            # 日付が15分以上前。警告のみ。
            if diff_time >= (old_m*60):
                print(f"WARNING: 入力元のICSファイル「{fname}」の日付が古いです。", file=sys.stderr)
                print(f"WARNING: {old_m}分以上前のファイルです。もし{old_m}分以内にダウンロードした場合は、", file=sys.stderr)
                print("WARNING: ファイル名が間違えてないか確認ください。", file=sys.stderr)

            #２回目は警告しない。
            F.old_file_check = False

        # utf-8-sig: BOM付きデータを読み込む。BOMがなければなにもしない。
        # https://docs.python.org/ja/2.7/library/codecs.html#module-encodings.utf_8_sig
        with open(fname, 'r', encoding='utf-8-sig') as f:
            ret = f.read()
        return ret
    # end of func

    #######################################
    @staticmethod
    def open_csv_object(fname: str):
        """
        出力先のファイルを開く。 文字列で"stdout"を指定すると標準出力となる。
    """
        if fname == "stdin"  or fname[0] == "-":
            raise RuntimeError(f"ファイル名指定エラー: {fname}")

        # errors='xmlcharrefreplace' utfからsjisに変換時にsjis未定義コードが出たときに "&#xxxx;"に変換する。
        # Ref: https://docs.python.org/ja/3/howto/unicode.html
        # Ref: https://zenn.dev/hassaku63/articles/f7ca587b86398c
        #
        #escale_type = 'backslashreplace'
        escale_type = 'xmlcharrefreplace'
        if fname == "stdout":
            #https://geroforce.hatenablog.com/entry/2018/12/05/114633
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, \
                                          encoding=F.CSV_ENCODING.name, \
                                          errors=escale_type, newline="")
            csv_out = sys.stdout
        else:
            if (not F.overwrite) and os.path.exists(fname):
                print(f"WARNING: CSVファイル 「{fname}」 がすでに存在します。")
                print("WARNING: 上書きしますか?")
                inp = input('WARNING: [Y]es/[N]o? >> ').lower()
                if not (inp in ('y', 'yes')):
                    print("WARNING: 処理を中断します。")
                    sys.exit(1)

            csv_out = open(fname, 'w', encoding=F.CSV_ENCODING.name, errors=escale_type, newline="")

        # Pythonライブラリの仕様でCSVの最後の改行はCR+LF。
        # Ref: https://docs.python.org/ja/3/library/csv.html
        # ->「Dialect.lineterminator」
        # -> 「writer が作り出す各行を終端する際に用いられる文字列です。デフォルトでは '\r\n' です。」

        return csv.writer(csv_out, quoting=csv.QUOTE_ALL)



class ModCSV:
    """CSVを加工する関係"""
    ##########################################################################
    #
    @staticmethod
    def set_summary_extend_head(mess: str, opt: str = None):
        """
        Summaryのheadを拡張します。
    """
        global F
        if opt is None:
            opt = ConstDat.SPLIT_SUMMAEY_EXTEND_HEAD_GENERIC

        for i in list(opt):
            if i.isspace():
                raise ValueError(f"ERROR: {mess}空白が含まれてます。")
            if not i.isprintable():
                raise ValueError(f"ERROR: {mess}使えない文字が含まれます。")
            if i in ConstDat.SPLIT_SUMMARY_HEAD_BAD_CHAR:
                raise ValueError(f"ERROR: {mess}使えない記号「{i}」が含まれます。")
        tmp_list = re.split('[,:]', opt)
        tmp_list2 = [item for item in tmp_list if item != '']

        if "Hidden" in tmp_list2:
                raise ValueError(f"ERROR: {mess}使えない文字列「Hidden」が含まれます。")

        F.SPLIT_SUMMARY_EXTEND_HEAD += tmp_list2
    #
    @staticmethod
    def split_garoon_style_summary(summary: str) -> str:
        """
    Garoonはタイトルは二種類の入力があり、
      タイトルの選択肢:'出張', '往訪', '来訪', '会議', '休み'
      タイトルの本文: 「東京特許許可局」

    などの入力がある。CSVではそれぞれ"予定","予定詳細"と出力される。

    ICSのSUMMARYにはこのような区別はないが、GaroonのCSVに変換するときにタ
    イトルの選択肢とコロンがあった場合は分割処理を行っている。

    例
    ICSの"SUMMARY":「出張:東京特許許可局」
    CSVの"予定":「出張」
    CSVの"予定詳細":「予定詳細」


        引数:
        summary: タイトルstr型で入った変数。

        返り値:
        summaryを加工後返す。

       外部制御変数:
        F.split_summary = True
        F.SPLIT_SUMMARY_EXTEND
    """
        # 以下正規表現文字列になるので変な記号は入れない
        # Garoonの選択肢のデフォルト。
        head = list(ConstDat.SPLIT_SUMMARY_HEAD)
        #SUMMARY分割の拡張
        head += F.SPLIT_SUMMARY_EXTEND_HEAD
        # コロンの半角。
        splitter = [':']
        # コロンの全角。
        splitter += ['：']

        for s in splitter:
            ret = summary.split(s, 1)
            if len(ret) == 1:
                continue
            ret[0] = ret[0].strip()
            ret[1] = ret[1].strip()
            for h in head:
                if re.fullmatch(h, ret[0]):
                    return ret[0], ret[1]
        #分割失敗
        return "", summary

    ###
    @staticmethod
    def modify_description(description: str) -> str:
        """
        メモ欄(description)の加工を行う。長いと見にくいのと、Teamsのパスワード
        が入ってることがあるので。

        引数:
        description: メモ欄がstr型で入った変数。

        返り値:
        descriptionを加工後返す。


       外部制御変数:
        F.description_delete_4th_line_onwards
        F.remove_teams_infomation

       Known bugs:
        Teamsの会議インフォメーションの削除は、フォーマットが変わったら無効です。
        2025/9現在のフォーマットをもとに削除を行います。

        """
        if description is None:
            return description

        if F.remove_teams_infomation:
            lines = description.splitlines()
            new_line = []
            for i in lines:
                if re.search("Microsoft Teams ヘルプが必要ですか", i):
                    new_line.append("(REMOVE TEAMS INFOMATION)")
                    break
                if re.search(r"\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.", i):
                    new_line.append("(REMOVE TEAMS INFOMATION)")
                    break
                if re.search("___________________________", i):
                    new_line.append("(REMOVE TEAMS INFOMATION)")
                    break
                new_line.append(i)
            description = "\n".join(new_line)
            description += "\n"

        if F.description_delete_4th_line_onwards:
            r = 4
            lines = description.splitlines()
            if len(lines) > r:
                description = "\n".join(lines[:r])
            else:
                description = "\n".join(lines)
            description += "\n"

        if F.remove_tail_cr:
            description = description.rstrip()
        return description

    ###
    # 登録番号記入の拡張仕様
    @staticmethod
    def enhanced_gyoumunum(description: str, summary: str) -> str:
        """Summary分割で、Summaryの最後尾に「-数字」もしくは「g数字」があった場合は、
        登録番号と見なし、DESCRIPTIONと置き換える。

        ※仕様検討中。

        スケジュールは技術部のみが使うものではなく他の教職員も使う。そのた
        め、謎の記述は極力避けるべきである。タイトル欄に謎の数字をいれるの
        は好ましくない。

        しかしながら、メモ欄にはTeams会議のパスワードなどセキュリティ情報がか
        かれる可能性があり、登録番号の摘出のためとはいえ、不必要に見える状
        態にするのは好ましくない。

    　　そのため、目立たない形で、タイトルの最後尾に記載する案とした。

        登録番号のための区切り文字の検討

        「&」「#」はxml生成時に別の意味が出るので不可
        「:」はGaroon形式に変換するときの「会議」「出張」とかの区切りと区別出来ないので不可
        「,」「;」はGaroonのICS生成時のバグの引き金になるので不可。
        「-」日付情報(12-31)やOS名(Windows-11)などを記載した場合誤認する
             それを許容したとしても、常に半角でいれてくれればいいが全角だ
             と似たようなのが多数あり面倒。全角ハイフン/全角マイナスなどい
             ろいろ。
        「$」はお金を書いてるのと区別つかないのと、一部プログラミング言語の変数
             表記なのでやめた方がよいかも。
        「/」は日付情報(12/31)などを記載した場合誤認する。また会議などの
             部屋番号等の記述でも見られる。
        「%」このあたりで試行する。
             誤用が無いとは言いきれないが。

        Known bugs:
        登録番号を4桁の数字としている。5桁以上なら要修正


        デバグコード:
          debug_modify_enhanced_gyoumunum.py

        処理の流れ:

        1 Summary: 文字列
          Summaryに登録番号がなければ、Descriptionは一切さわらない。無変。

        2. Summary: abcd %数字A or
           Summary: abcd g数字A

        SUMMARYに登録番号があった場合は、SUMMARY側が優先されて、
        DESCRIPTIONの変換を行う。

        Description 変換規則
         方針1:行数の変化は可能な限り避ける。
         方針2:1行めにある登録番号とSUMMARYの登録番号が異なれば置き換える。
         方針3:1行めに空行があれば登録番号と置き換える。
         方針4:1行目に登録番号と空行以外がある場合は、行数を増やす。
         方針5:1行めに「可」「急」があれば、先頭に空行を1行追加して、1行めに登録番号を書き込む
         方針6:1行めに「可」「急」以外があれば、先頭に空行2行追加して、1行めに登録番号を書き込む


        以下
         Pre: DESCRIPTION変換前
         Aft: DESCRIPTION変換後
         数字A: SUMMARYに記載があった登録番号
         数字B: DESCRIPTIONに記載があった登録番号

        "(N/A)": CHANGELOG.mdにも記載したが、DESCRIPTIONが未定義という事をしめす
             特殊な文字。改行のみがあった場合は未定義ではなく""が入る。

        (空白文字): TAB, 全角SPACE (改行は含まない。)
        (改行文字): \\n  (改行の正規化をしているので\\rは出ないはず。)
        「\\s」: 改行を含む空白文字(SPACE, TAB, 全角SPACE, \\n)
        「.」: 任意の一文字

        Description-Type1-1:
        注: 行数が減る可能正あり。
        -(Pre)-------
        (空白文字)*数字B\\s*
        -------------
        -(Aft)-------
        数字A
        -------------

        Description-Type1-2:
        注:"(N/A)"の後ろの文字は削除。
        注: 行数が減る可能正あり。
        -(Pre)-------
        (N/A).*
        -------------
        -(Aft)-------
        数字A
        -------------

        Description-Type1-3:
        注: 行数が減る可能正あり。
        -(Pre)-------
        \\s*
        -------------
        -(Aft)-------
        数字A
        -------------

        Description-Type2:
        行数は変化しない。
        -(Pre)-------
        (空白文字)*数字B(空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        数字A(改行文字).*
        -------------

        Description-Type3:
        行数は変化しない。
        -(Pre)-------
        (空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        数字A*(改行文字).*
        -------------

        Description-Type4:
        行数が1行増える
        -(Pre)-------
        [可|急](空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        (空白文字)*数字A(改行文字)[可|急](空白文字)*(改行文字).*
        -------------

        Description-Type5:
        行数が2行増える
        -(Pre)-------
        .*
        -------------
        -(Aft)-------
        数字A(改行文字)(改行文字).*
        -------------

        """
        # 登録番号記入の拡張仕様: SUMMARYの「g」と「%」
        m = re.search(r"[ｇg%％]([0-9０-９]{1,4})[　 \t]*$", summary)
        if m is None:
            return None
        # 全角数字を半角にするため、0を足してる。
        #print(m.groups())
        #print(m.group())
        gyoumunum = str(int(m.groups()[0])+0)

        if (int(gyoumunum) < 0) or (int(gyoumunum) > 9999):
            # 負の数は登録番号としては無効
            # 5桁の登録番号は無効(正規表現的にないはずだが。)
            raise RuntimeError("ERROR: Summaryの登録番号の取得に失敗しました")

        if re.search(r"\r", description, flags=re.DOTALL):
            raise RuntimeError("ERROR: 改行の正規化が行われてません。「\\n」のみ有効です。")

        # re.matchは1行めのみ検索する。2行目は見ない。
        #print(f"DEBUG: found enhance gyoumunum = {gyoumunum}")

        # Description-Type1-1:
        # 正規表現の空白のところに全角スペース入ってる
        # 改行の正規化をしているので改行に「\r」は出ないはず。
        m1 = re.search(r"^[　 \t]*[0-9０-９]{1,4}[　 \t\n]*$", description, flags=re.DOTALL)

        # Description-Type1-2:
        # 注:"(N/A)"の後ろの文字は無視。
        if ConstDat.NA != '(N/A)':
            raise RuntimeError("ERROR: Override(N/A)")
        # 仕様変更で、現時点ではNoneが入っていると思われるので、
        # 実際はこのコードは動かない可能性が高い。
        m2 = re.search(r"^\(N\/A\).*$", description, flags=re.DOTALL)

        # Description-Type1-3:
        # 空白と改行のみ。
        m3 = re.search(r"^[　 \t\r\n]*$", description, flags=re.DOTALL)
        if m1 or m2 or m3:
            return gyoumunum

        # Description-Type2:
        # 行数は変化しない。
        m1 = re.search(r"^[　 \t]*[0-9０-９]{1,4}[　 \t]*\n.*", description, flags=re.DOTALL)
        if m1:
            ret = re.sub(r"^[　 \t]*[0-9０-９]{1,4}[　 \t]*", gyoumunum, description)
            if ret is None:
                raise RuntimeError("ERROR: 正規表現の想定外のエラー")
            return ret

        # Description-Type3:
        # 行数は変化しない。
        m1 = re.search(r"^[　 \t]*\n.*", description, flags=re.DOTALL)
        if m1:
            return re.sub(r"^[　 \t]*", gyoumunum, description)
        # Description-Type4:
        # 冒頭が「可急」の場合は登録番号を行頭に差し込む。行数が1行増える。
        lines = description.splitlines()
        m1 = re.search(r"^[　 \t]*[可急][　 \t]*$", lines[0])
        if m1:
            if re.search("可", lines[0]):
                lines[0] = "可"
            if re.search("急", lines[0]):
                lines[0] = "急"
            return gyoumunum + "\n" + "\n".join(lines) + "\n"

        # Description-Type5:
        #それ以外は登録番号を行頭に差し込み改行を2個差し込む。行数が2行増える。
        return gyoumunum + "\n\n" + description
    ###

    @staticmethod
    def modify_csv(csv_buffer: list, timerange: int) -> list:
        """
        1. timerange範囲外のデータをすてる
        2. 各種加工を行う。出力対象のCSVの行数をlistで返す。
        """
        ###################
        ret_index = []
        # 無効なデータを捨てながら各種加工を行う。
        for i in range(len(csv_buffer)):
            # 範囲外/無効なデータを捨てる。

            if csv_buffer[i][F.CSV_POS2["H:UID"]] is None:
                continue
            if csv_buffer[i][F.CSV_POS2["H:DTSTART"]] is None:
                continue
            if not TimeRange.is_collect(csv_buffer[i][F.CSV_POS2["H:DTSTART"]], timerange):
                continue

            ret_index.append(i)

            # ICSのデータで指定の要素がなかった場合はNoneが入っている。
            # 適切な用語に書き換える。
            for j in range(F.CSV_POS2["H:LENGTH"], len(F.CSV_HEADER)):
                if csv_buffer[i][j] is None:
                    csv_buffer[i][j] = ConstDat.NA
                elif F.remove_tail_cr:
                    #各要素の最後の改行と空白をすべて取り除く。
                    csv_buffer[i][j] = csv_buffer[i][j].rstrip()

            summary = csv_buffer[i][F.CSV_POS2["SUMMARY"]]

            if F.split_summary and (not summary is None):
                #ICS形式の場合は予定(選択肢のところ)が無いので生成試みる
                summary_h, summary = ModCSV.split_garoon_style_summary(summary)
                csv_buffer[i][F.CSV_POS2["SUMMARY:H"]] = summary_h
                csv_buffer[i][F.CSV_POS2["SUMMARY"]] = summary

            if 'DESCRIPTION' not in F.CSV_POS:
                continue

            description = csv_buffer[i][F.CSV_POS2["DESCRIPTION"]]
            description = ModCSV.modify_description(description)

            # 登録番号記入の拡張仕様
            # SUMMARYに記載された登録番号をDESCRIPTIONに差し込む。
            if F.enhanced_gyoumunum:
                d = ModCSV.enhanced_gyoumunum(description, summary)
                if d:
                    description = d
            csv_buffer[i][F.CSV_POS2["DESCRIPTION"]] = description
        # end for i
        return ret_index

class RecurrenceID:
    """RecurrenceID関連処理"""
    @staticmethod
    def id_list_dump(l: dict, prefix="DEBUG:", file=sys.stderr):
        """
        recurrence_id_listをdump
    """
        print("----", file=file)
        for uuid in l.keys():
            print(f'{prefix}uid = {uuid}', file=file)
            for vv in l[uuid]:
                print(f'{prefix}\tRECURRENCE-ID = {vv}', file=file)
        print("----", file=file)

    #End of func()
    @staticmethod
    def restore_aux(buff: list, recurrence_id_list: dict, outlook_bugfix=False) -> int:
        """
        RECURRENCE-IDで上書きするVEVENTはSUMMARYやDESCRIPTIONが未定義
        の場合がある。本関数で、復元する。

        なお、buffを参照渡しで使ってるため、buff[x] の値を直接書き換えずに、
        buff[x][y]を書き換えてください。

        空文字の場合と未定義の場合を区別するため、未定義の場合は"(N/A)"が入ってる。

        restore_aux()を呼び出したあと、復元に失敗しているVEVENTがあれば、
        outlook_bugfix=Trueにして再度呼び出している。デバック時は注意。

        """
        bad_count = 0

        uid2line = {}
        for i in range(len(buff)):
            uid = buff[i][F.CSV_POS2["H:UID"]]
            if uid in recurrence_id_list:
                if not uid in uid2line:
                    uid2line[uid] = []
                uid2line[uid].append(i)

        key_list = list(recurrence_id_list.keys())
        for key in key_list:
            line_list = uid2line[key]
            for i in line_list:
                b = buff[i]
                uid = b[F.CSV_POS2["H:UID"]]
                recurrence_id = b[F.CSV_POS2["H:RECURRENCE_ID"]]

                if recurrence_id is None:
                    continue

                flag_found_j = -1
                for j in line_list:
                    if i == j:
                        continue
                    if not buff[j][F.CSV_POS2["H:RECURRENCE_ID"]] is None:
                        continue

                    if recurrence_id == buff[j][F.CSV_POS2["H:DTSTART"]]:
                        flag_found_j = j
                        break

                    # 修正前のdtstartがdatetime.date(日付のみ)なのに、
                    # 修正先のrecurrence_idがdatetime.datetime(日時情報あり)になっとる。
                    if outlook_bugfix:
                        dd = TZ.naive2aware(buff[j][F.CSV_POS2["H:DTSTART"]])
                        if recurrence_id == dd:
                            flag_found_j = j
                            break

                k = -1
                if flag_found_j < 0:
                    bad_count = bad_count + 1
                    k = F.CSV_POS2["SUMMARY"]
                    if buff[i][k] is None:
                        buff[i][k] = ConstDat.UNREF
                    continue

                recurrence_id_list[uid].remove(recurrence_id)
                if len(recurrence_id_list[uid]) == 0:
                    del recurrence_id_list[uid]

                for k in range(F.CSV_POS2["H:LENGTH"], len(F.CSV_HEADER)):
                    if (buff[i][k] is None) or (outlook_bugfix and buff[i][k] == ConstDat.UNREF):
                        buff[i][k] = buff[flag_found_j][k]

                if F.override_recurrence_id:
                    buff[flag_found_j][F.CSV_POS2["H:UID"]] = None
                else:
                    k = F.CSV_POS2["SUMMARY"]
                    buff[flag_found_j][F.CSV_POS2["H:UID"]] += "Hidden"
                    buff[flag_found_j][k] = "Hidden: " + buff[flag_found_j][k]

                #end for j
            #end for i
        #end for key
        return bad_count

    # end of func()
    @staticmethod
    def restore(buff: list, recurrence_id_list: dict) -> int:
        """
        RECURRENCE-IDで上書きするVEVENTはSUMMARYやDESCRIPTIONが未定義の場合が
        ある。復元する。
    """
        bad_recurrence_id_count = RecurrenceID.restore_aux(buff, recurrence_id_list)
        if bad_recurrence_id_count > 0:
            bad_recurrence_id_count =  \
                RecurrenceID.restore_aux(buff, recurrence_id_list, outlook_bugfix=True)
            if bad_recurrence_id_count > 0:
                print(f"WARNING: 繰返し命令の{bad_recurrence_id_count}\
                        個の復元に失敗。", file=sys.stderr)
                print("WARNING: 失敗したスケジュールは\
                     「(REFERENCE DATA DOES NOT EXIST)」\
                     と記載されています。", file=sys.stderr)
                print("WARNING: 手作業で修正してください。", file=sys.stderr)
                if not F.DEBUG_UID is None:
                    print(f"D2: recurrence_id_list = {recurrence_id_list}", file=sys.stderr)

        # 浮いてたrecurrence_id_list。
        # 例えば4月のスケジュール生成時に3月のデータを引用してたら浮く。
        # Outlook(classic)の場合は「(REFERENCE DATA DOES NOT EXIST)」となる。
        # Outlook(Web)の場合はこの問題は発生しない。
        if len(recurrence_id_list) > 0:
            print("WARNING: ICSファイルの異常です。参照先が無い繰返し命令があります。", file=sys.stderr)
            RecurrenceID.id_list_dump(recurrence_id_list, "BROKEN VEVENT:")

        return bad_recurrence_id_count

class Main:
    """ICSからCSVに変換する関数の親の関数"""
    @staticmethod
    def ics_parts_to_csv_buffer(ics_parts, rrule_start=None) -> list:
        """
       VEVENTをCSV出力用の文字列(list)に変換する。

        引数:
        ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
                   入ってる。

        rrule_start: 繰返しスケジュールの時の開始時刻。VEVENTのDTSTARTとDTENDを置き換える。
                     Noneの場合は、VEVENTのDTSTARTとDTENDがそのまま使われる。
                     datetime.datetime型もしくはdatetime.date型

        返り値: CSV出力用の文字列に変換してLISTにいれて返す。

       外部制御変数:
        F.remove_tail_cr

        独自定義のICS要素の追加手順:

        ICSファイルの要素ABCの中身を加工せずにそのまま出力する場合は
        PreSetup.set_format()で
           F.CSV_POS["ABC"]  = CSVの位置
        と記載します。

        何らかの加工をする場合は「:」付きで以下のような感じで記載し
           F.CSV_POS["X:ABC"]  = CSVの位置
           F.CSV_POS["ABC:適当な名前"]  = CSVの位置
        以下の関数に処理手順を記載します。

    """
        p = F.CSV_POS.copy()
        row = [None] * F.CSV_POS2["B:LENGTH"]

        # 特殊処理を必要とするICS要素
        s_d, s_t, e_d, e_t, allday = TZ.ics_parts_to_csv_time(ics_parts, rrule_start)

        # 必須項目とする。
        row[p.pop("DTSTART:DAY")] = s_d
        row[p.pop("DTSTART:TIME")] = s_t
        row[p.pop("DTEND:DAY")] = e_d
        row[p.pop("DTEND:TIME")] = e_t

        if 'X:ALLDAY_EVENT' in F.CSV_POS:
            row[p.pop("X:ALLDAY_EVENT")] = allday

        # 任意項目とする
        if 'SUMMARY:H' in p:
            row[p.pop("SUMMARY:H")] = ""

        #F.CSV_POS["X-MICROSOFT-CDO-BUSYSTATUS:NUM"] = 15
        if 'X-MICROSOFT-CDO-BUSYSTATUS:NUM' in p:
            pos = p.pop('X-MICROSOFT-CDO-BUSYSTATUS:NUM')
            n = Misc.get_ics_val(ics_parts, 'X-MICROSOFT-CDO-BUSYSTATUS', None, exit_none=False)
            if not n is None:
                row[pos] = ConstDat.CSV_TABLE_X_MICROSOFT_CDO_BUSYSTATUS[n]


        #F.CSV_POS["ORGANIZER:CN"] = 9
        if 'ORGANIZER:CN' in p:
            pos = p.pop('ORGANIZER:CN')
            n = Misc.get_ics_val(ics_parts, 'ORGANIZER', None, exit_none=False)
            if not n is None:
                if 'CN' in ics_parts.organizer.params:
                    row[pos] = ics_parts.organizer.params['CN'][0]

        if 'ATTENDEE:CN:RSVP:TRUE' in p:
            pos = p.pop('ATTENDEE:CN:RSVP:TRUE')
            n = Misc.get_ics_val(ics_parts, 'ATTENDEE', None, exit_none=False)
            if not n is None:
                #print(ics_parts.attendee.params)
                cn = []
                for user in ics_parts.attendee_list:
                    if ('CN' in user.params) and ('RSVP' in user.params):
                        if user.params['RSVP'][0].lower() == 'true':
                            cn.append(user.params['CN'][0])
                row[pos] = ";".join(cn)

        if 'ATTENDEE:CN:RSVP:FALSE' in p:
            pos = p.pop('ATTENDEE:CN:RSVP:FALSE')
            n = Misc.get_ics_val(ics_parts, 'ATTENDEE', None, exit_none=False)
            if not n is None:
                #print(ics_parts.attendee.params)
                cn = []
                for user in ics_parts.attendee_list:
                    if ('CN' in user.params) and ('RSVP' in user.params):
                        if user.params['RSVP'][0].lower() == 'false':
                            cn.append(user.params['CN'][0])
                row[pos] = ";".join(cn)


        # 特殊処理を必要としないICS要素
        for k, v in p.items():
            if ':' in list(k):
                raise ValueError(f"Internal Error: 未対応のICS要素: {k}")
            row[v] = Misc.get_ics_val(ics_parts, k, None, exit_none=False)

        return row

    ###
    @staticmethod
    def vobject2csv(calendar: vobject.base.Component, timerange: int):
        """
        補助関数。 vobjectを読み込んで、csv出力用のbufferにいれていく。


        引数timerangeは使ってないが、念のため残しておく。
    """
        # 返り値
        # VERSION1.3追加
        # 旧版1.2ではCSVの要素を生成したらすぐ出力してたが、
        # 上書スケジュール(RECURRENCE-ID)対応のためバッファリングを行う。
        csv_buffer = []

        # key: UID, value: RECURRENCE-IDをリストで収納。
        # 処理段階で基になるVEVENTと読み替え成功したら、消していく。
        # 最終的に残るのが読み替えに失敗したRECURRENCE-ID。
        recurrence_id_list = {}

        for component in calendar.components():
            if component.name != 'VEVENT':
                continue

            dtstart = Misc.get_ics_val(component, 'dtstart')
            dtend = Misc.get_ics_val(component, 'dtend')
            uid = Misc.get_ics_val(component, 'uid', ConstDat.NA)
            rrule = Misc.get_ics_val(component, 'rrule', None, exit_none=False)
            # VERSION1.3追加: RECURRENCE-IDコード。
            recurrence_id = Misc.get_ics_val(component, 'recurrence-id', None, exit_none=False)
            recurrence_id = TZ.to_localtime(recurrence_id, exit_none=False)

            # debugコード
            if (F.DEBUG_UID is not None) and F.DEBUG_UID != uid:
                continue

            # データの検査
            if (not rrule is None) and (not recurrence_id is None):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにRECURRENCE-IDとRRULEがあります。")

            # Known bugs: RDATEに対応はしたが、動作確認例が少ないため、要注意。
            #if Misc.get_ics_val(component, 'rdate', ConstDat.NA) != ConstDat.NA:
            #    raise RuntimeError("ERROR: 本プログラム未実装のICS命令RDATEが使われています。")
            #
            if TZ.is_aware(dtstart) != TZ.is_aware(dtend):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにdtstart/dtendにtimezone有り/無しが混在。")

            if TZ.hava_time(dtstart) != TZ.hava_time(dtend):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにdtstart/dtendに時刻情報の有り/無しが混在。")

            if F.DEBUG_UID == uid:
                print(f"STEP1: uid = {uid}", file=sys.stderr)
                print(f"STEP1: dtstart = {dtstart}", file=sys.stderr)
                print(f"STEP1: dtend = {dtend}", file=sys.stderr)

            # CSV用のlist生成開始。
            buff_pre = [uid, TZ.to_localtime(dtstart), recurrence_id]
            buff_aft = Main.ics_parts_to_csv_buffer(component)

            # ICSのRRULE命令が未使用ならそのまま出力する。
            if rrule is None:
                csv_buffer.append(buff_pre + buff_aft)

                if F.support_recurrence_id and (not recurrence_id is None):
                    if uid not in recurrence_id_list:
                        recurrence_id_list[uid] = []
                    recurrence_id_list[uid].append(recurrence_id)
                continue

            #ICSのRRULE命令の処理。
            #t_c = component
            #if type(dtstart) is datetime.date:
            #print(type(rrule))
            #print(rrule)
            #ignoretz = (not isinstance(dtstart, datetime.datetime) or dtstart.tzinfo is None)

            # rruleの繰返し回数はcountで指定と最終日時のuntilの場合がある。
            # untilの場合はいろいろ大変。バグが非常に出やすい。

            # Known bugs: 内部変数「_until」にアクセスしているため、ライブラリの仕様が
            # 変わったら動かない。
            until = dateutil.rrule.rrulestr(rrule)._until

            org_dtstart = dtstart
            if not until is None:
                if TZ.is_aware(dtstart) and TZ.is_naive(until):
                    # dtstartがaware(timezone有)であり、untilがnaive(floatingtime)。
                    raise ValueError("ERROR: ICSデータ不整合: \
                                     同一VEVENTのdtstartはtimezone有りで、\
                                     rruleのuntilがtimezone有無し。")

                if F.naive_aware_mixed_bugfix and TZ.is_aware(until) and TZ.is_naive(dtstart):
                    # dtstartがnaive(floatingtime)だが、untilがaware(timezoneあり)の場合。
                    # 本来はICSデータの不整合なのだが、あまりにこの事例が多いため対処。
                    component.add('x-org-dtstart').value = dtstart
                    component.dtstart.value = TZ.naive2aware(dtstart)
                    component.add('x-org-dtend').value = dtend
                    component.dtend.value = TZ.naive2aware(dtend)

                    if F.DEBUG_UID == uid:
                        t = Misc.get_ics_val(component, 'dtstart')
                        print(f"STEP2: aware dtstart  = {t}", file=sys.stderr)
                        t = Misc.get_ics_val(component, 'dtend')
                        print(f"STEP2: aware dtend  = {t}", file=sys.stderr)

            #component.prettyPrint()

            rrule_set = component.getrruleset(addRDate=True)
            if F.DEBUG_UID == uid:
                print(f"STEP2.5: RRULE = {rrule}", file=sys.stderr)
                for s in rrule_set:
                    print(f"RRULE_PARTS={s}", file=sys.stderr)

            for s in rrule_set:
                # getrrulesetがdatetime.dateからdatetime.datetimeに拡張する事がある。
                if TZ.hava_time(s) and (not TZ.hava_time(org_dtstart)):
                    if not TZ.is_am12(s):
                        raise ValueError("BUG: getrrulesetの計算がおかしい")
                    s = s.date()


                if F.DEBUG_UID == uid:
                    print(f"STEP3: s   = {s}", file=sys.stderr)

                if F.DEBUG_UID == uid:
                    print(f"STEP4: PASS(timerange={timerange})", file=sys.stderr)

                t = TZ.ics_parts_to_csv_time(component, s)
                buff_pre[F.CSV_POS2["H:DTSTART"]] = TZ.to_localtime(s)
                #buff_aft[0:4] = t
                buff_aft[F.CSV_POS["DTSTART:DAY"]] = t[0]
                buff_aft[F.CSV_POS["DTSTART:TIME"]] = t[1]
                buff_aft[F.CSV_POS["DTEND:DAY"]] = t[2]
                buff_aft[F.CSV_POS["DTEND:TIME"]] = t[3]
                if 'X:ALLDAY_EVENT' in F.CSV_POS:
                    buff_aft[F.CSV_POS["X:ALLDAY_EVENT"]] = t[4]

                csv_buffer.append(buff_pre + buff_aft)
                if F.DEBUG_UID == uid:
                    print(f"STEP4: normailize(s) = \
                           {buff_pre[F.CSV_POS2['H:DTSTART']]}", file=sys.stderr)
        # end for()
        return csv_buffer, recurrence_id_list
    #end of func.


    #####
    @staticmethod
    def ics2csv(ics_file_path: str, csv_file_path: str, timerange: int = 0) -> None:
        """
        ICS(iCalendar)ファイルをCSVファイルに変換する。

        引数:
            ics_file_path (str): 変換元のICS(iCalendar)ファイル。"stdin"を指定すると標準入力。
            csv_file_path (str): 変換先のCSVファイル。"stdout"を指定すると標準出力。
            timerange (int): CSVに変換する日時を限定する場合は、指定する。
                             2025年8月分がほしい場合は「202508」と指定する。
                             未指定や「0」だと全部変換する。
        返り値:
            None。失敗したら停止する。
        """
        ######################
        if vobject.VERSION != "0.9.9":
            print("ERROR: 依存ライブラリvobjectのバージョンが開発環境と異なります。", \
                  file=sys.stderr)
            print(f"ERROR: githubよりプルリクエストください。 vobject.VERSION={vobject.VERSION}", \
                  file=sys.stderr)
            sys.exit(1)

        ########
        #
        if False:
            print(f"INFO: CSVフォーマット: {F.CSV_FORMAT}", file=sys.stderr)
            print(f"INFO: 文字コード: {F.CSV_ENCODING}", file=sys.stderr)
            print(f"INFO: AllDayFormat: {F.CSV_ALLDAY_FORMAT}", file=sys.stderr)
            print(f"INFO: CSVヘッダ: {F.CSV_HEADER}", file=sys.stderr)

        ######################
        #ファイルから読み込み
        ics_data = FileIO.file2str(ics_file_path)
        # あまりに小さい。
        if len(ics_data) < 10:
            raise RuntimeError(f"ERROR: ファイル読み込みエラー: ファイル行数: {len(ics_data)}")

        ######################
        # ライブラリvobjectがのicsファイルの読み込む時に例外を履く
        # 記述の修正を行う。
        # liics2gacsv(v1.4)では RRULEのEXDATEの記述の修正のみ。
        if F.exdate_format_bugfix:
            ics_data = PreSetup.bugfix_exdate_format(ics_data)
        else: # 処理の都合で改行の正規化が必要
            ics_data = "\n".join(ics_data.splitlines()) + "\n"

        ######################
        # 読み込んだデータstrをvobjectに変換。
        calendar = vobject.readOne(ics_data)

        ######################
        # TimeZoneデータ読み込み
        TZ.load_ics(ics_data, F.OVERRIDE_TIMEZONE)

        ######################
        # vobjectのオブジェクトをCSVに変換
        csv_buffer, recurrence_id_list = Main.vobject2csv(calendar, timerange)

        Misc.csv_buffer_dump(csv_buffer, prefix="D1:", uid=F.DEBUG_UID)

        ######################
        # RECURRENCE-IDで上書きするVEVENTはSUMMARYやDESCRIPTIONが未定義の場合が
        # ある。復元する。
        # RECURRENCE-IDで上書きするVEVENTを消す。もしくは印をつける。

        bad_recurrence_id_count = 0
        if F.support_recurrence_id:
            bad_recurrence_id_count = RecurrenceID.restore(csv_buffer, recurrence_id_list)

        if not F.DEBUG_UID is None:
            Misc.csv_buffer_dump(csv_buffer, prefix="D2:", uid=F.DEBUG_UID)
            print(f"D2: recurrence_id_list = {recurrence_id_list}", file=sys.stderr)

        ######################
        #timerange範囲外のデータを捨てる。
        # 各種加工を行う。出力対象のCSVの行数をlistで返す
        csv_index = ModCSV.modify_csv(csv_buffer, timerange)

        ######################
        # 日付でsortする. index sort.
        # BUG: 月と日が2桁でないとバグる。 可:2026/01/02, 不可:2026/1/2
        # BUG: 順番が (「20xx/月/日」| ISO_8601)以外だったら不可
        if F.output_sort:
            k = [F.CSV_POS2["DTSTART:DAY"], F.CSV_POS2["DTSTART:TIME"], \
                 F.CSV_POS2["DTEND:DAY"], F.CSV_POS2["DTEND:TIME"], \
                 F.CSV_POS2["SUMMARY"]]

            csv_index.sort(key=lambda x: [csv_buffer[x][k[0]], csv_buffer[x][k[1]],\
                                          csv_buffer[x][k[2]], csv_buffer[x][k[3]],\
                                          csv_buffer[x][k[4]]])

        Misc.csv_buffer_dump(csv_buffer, prefix="D3:", uid=F.DEBUG_UID)

        # 出力用CSVファイルのopen。
        csv_writer = FileIO.open_csv_object(csv_file_path)

        #CSVのHeader出力
        if F.print_csv_header:
            csv_writer.writerow(F.CSV_HEADER[F.CSV_POS2["H:LENGTH"]:])

        #CSVの要素出力
        for i in csv_index:
            csv_writer.writerow(csv_buffer[i][F.CSV_POS2["H:LENGTH"]:])

        Misc.csv_buffer_dump(csv_buffer, prefix="D4:", uid=F.DEBUG_UID)

        # 終了ステータス表示。
        if bad_recurrence_id_count == 0:
            print(f"INFO: 変換に成功しました: '{ics_file_path}' to '{csv_file_path}'",\
                  file=sys.stderr)
        else:
            print(f"WARNING: 変換に*概ね*成功しました: '{ics_file_path}' to '{csv_file_path}'",\
                  file=sys.stderr)

    #end func


############################################
HELP_PART1 = f"""
必須引数:

``期間''

CSVの出力年月の期間指定。有効範囲は2000年1月から2099年12月まで。

西暦4桁+月2桁の合計6桁の数字で指定します。

例えばCSVに2025年8月分を出力する場合は「202508」と指定する。

特殊な値として、「all」もしくは「0」だと全期間を変換する。
「guessin」だと入力ファイル名から期間を推測する。
「guess」だと出力ファイル名から期間を推測する。例えば出力ファイル名が
「schedules202509.csv」なら2025年9月と推測する。

期間指定を行った場合、月末のスケジュールで月を超えてる場合は翌月分も
含まれます。例えば11月30日23:00に開始で12月1日02:00に終了の場合は、11
月分に12月1日02:00終了のスケジュールが入ります。12月分には入りません。

``入力.ics''

変換元のICSファイル名を指定。「stdin」を指定すると標準入力。ICSファ
イルは規格(RFC5545)で文字コードがUTF-8と決まってます。そのため、必ず
UTF-8のICSファイルを指定してください。

  ※対応ICSファイル:
  Cybozu Garoon(Version 5.0.2)
  Web版 Outlook
  Windowsアプリ版 Outlook(classic)
  いずれも2025年9月から12月ごろに生成されたICSファイルの出力で確認。

``出力.csv''

変換先のCSVファイル名を指定。「stdout」を指定すると標準出力。

=====================================================================
オプション引数:

* ヘルプ:

-h, --help
ヘルプを出力する。ヘルプ画面の停止はアルファベットの「q」を押してくだ
さい。

* CSVのヘッダ関係:

-k, --print-csv-header
CSVの一行めに項目一覧を表示する。defaultは表示しない。

タイトル(SUMMARY)の分割関係:
※CSVがGaroon形式の場合のみ有効です。

-s, --disable-split-summary
ICSのSUMMARYの分割を無効にする。CSVがGaroon形式の場合はdefaultは有効で
あり分割します。分割するとCSVの「予定」/「予定詳細」がGaroonと同じ形式
になる。

※詳細は関数ModCSV.split_garoon_style_summary()をみよ。

作者の職場向けの拡張機能:

-m, --extend-summary-head
ICSのSUMMARYの分割で拡張。defaultは無効。

デフォルトで有効な予定:
  出張,往訪,来訪,会議,休み

引数を指定すると下記が追加:
  {ConstDat.SPLIT_SUMMAEY_EXTEND_HEAD_GENERIC}

※詳細は関数ModCSV.split_garoon_style_summary()とPreSetup.parse_args()
をみよ。

--add-summary-head="文字列,文字列", --add-summary-head="文字列:文字列"
ICSのSUMMARYの分割でヘッダを追加します。複数指定できます。「,」もしく
は「:」で区切ってください。文字列に記号の指定は不可。

例
  --add-summary-head="研究,教育" :
  --add-summary-head="支部:本部" :

※内部処理の都合で、引数に「Hidden」は指定できません。繰返しスケジュー
ルの一部上書き(RECURRENCE_ID)で別用途で使っているため。

-z, --enhance-tourokunum, --enhance-touroku-number
--enhance-gyoumunum, --enhance-gyoumu-number
SUMMARYの最後尾に「%数字」もしくは「g数字」があった場合は登録番号(業務
番号)と見なし、メモ欄(description)に登録番号を書き込む。defaultは無効。

メモ欄(description)に最初から登録番号と考えられる数字の記載があっ
た場合はSUMMARYに記載された登録番号を優先し、メモ欄(description)
の登録番号を書き換える。
※v2.1で追加。
※仕様検討中。
※詳細は関数ModCSV.enhanced_gyoumunum()をみよ。
"""

HELP_PART2 = """
* CSVの文字コードの指定

-C"文字列", -Cshift_jis, -Cutf_8, -Cutf_8_sig
CSV出力の文字コードを指定する。

defaultはutf-8。ただし後述のCSVの書式指定オプション「-F"文字列"」で
defaultが変更になります。

* CSVの書式を指定

-F"文字列"
CSVの書式を指定します。但しICS生成時に出力内容を制限した場合、ICSから取得
可能な要素のみ出力を行います。

-FSimple
Default. 文字コードはUTF-8。時刻形式はiso8601(extended)

概ねICSの要素のベタ展開。将来CSV要素を後ろに追加する形で増やす可能性あ
ります。

  CSVヘッダ: "DTSTART:DAY","DTSTART:TIME", "DTEND:DAY", "DTEND:TIME",
  "SUMMARY" , "DESCRIPTION", "X-MICROSOFT-CDO-BUSYSTATUS", "CATEGORIES"

  ただし、"X-MICROSOFT-CDO-BUSYSTATUS"はOutlook(Web, classic)のICSのみ
  対応。"CATEGORIES"はOutlook(Web, classic)のICSのみ対応。それ以外の
  ICSでは"(N/A)"という出力になります。

-FGaroon
Garoonとほぼ同等の出力を行います。文字コードはShift_JIS

  CSVヘッダ: "開始日", "開始時刻", "終了日", "終了時刻", "予定",
             "予定詳細", "メモ"

-FOutlookClassic (未実装)
OutlookClassicとほぼ同等の出力を行います。文字コードはUTF-8。

  CSVヘッダ: "件名","開始日","開始時刻","終了日","終了時刻",
             "終日イベント","アラーム オン/オフ","アラーム日付",
             "アラーム時刻","会議の開催者","必須出席者","任意出席者",
             "リソース","プライベート","経費情報","公開する時間帯の種類",
              "支払い条件","場所","内容","秘密度","分類","優先度"

  ただし、有効な要素は
         "件名","開始日","開始時刻","終了日","終了時刻",
          "終日イベント","会議の開催者","必須出席者","任意出席者"

 のみです。それ以外は"(N/A)"という出力になります。会議の出席者等は
 Outlook(classic)のICSのみ対応。

* ICSのTimeZone指定

-T"文字列"

一部のICSファイルはTimeZoneの定義が複数行われている場合がある。採用す
るTimeZoneを指定する。日本標準時なら -T"Asia/Tokyo"などを定義すればよ
い。

※本オプションをWindows環境で使う場合はライブラリ「tzdata」をインストー
ルしてください。

※TimeZoneが不適切な場合、日時計算に失敗し、誤ったCSVが生成されます。

=====================================================================

特殊なオプション引数:

以降で述べる引数は多くの場合は指定する必要は無いです。

* TimeZoneの表記関係

スケジュールで「終日スケジュール」関連は取扱いに細心の注意が必要です。
詳細 misc/TECH-MEMO.txt に詳細に記載してます。

--show-timezone
CSVの出力時刻にTimeZone情報を表示する。ICSファイルにTimeZone情報が含ま
れていた場合、TimeZoneを表示する。defaultでは表示しない。

  default(もしくはTimeZone情報なし):
    12:34:55
    12:34:55

  オプション「--show-timezone」指定(TimeZone情報あり):
    12:34:55+0900
    12:34:55-0300

※TimeZoneが複数定義が行われてる場合は正常に動作しない可能性あり。
※Garoonが生成するICSファイルにはTimeZone情報が無い。
※詳細は関数TZ.ics_parts_to_csv_time()をみよ。

* 終日スケジュールの表記関係
ICSでは終日スケジュールで時刻がある「0:00開始翌日0:00終了」の場合と、
時刻がない終日スケジュールがあり、区別される。CSVへの出力形式を指定す
る。

※詳細は関数TZ.ics_parts_to_csv_time()および misc/TECH-MEMO.txt をみよ。

--allday-format-today
終日スケジュール(時刻なし)の出力形式を「時刻なしの、当日開始、当日終了」
とする。引数「-FGaroon」指定時は本オプションがdefaultになります。

オプション「--allday-format-today」指定時の終日スケジュール(時刻なし)
の出力:
  "2025/09/05","","2025/09/05","" (24時間の終日スケジュール)
  "2025/09/05","","2025/09/06","" (48時間の終日スケジュール)

--allday-format-nextday
終日スケジュール(時刻なし)の出力形式を「時刻なしの、当日開始、翌日終了」
とする。引数「-Fsimple」指定時は本オプションがdefaultになります。

オプション「--allday-format-nextday」指定時の終日スケジュール(時刻な
し)の出力:
  "2025/09/05","","2025/09/06","" (24時間の終日スケジュール)
  "2025/09/05","","2025/09/07","" (48時間の終日スケジュール)

※ICSの内部形式は本形式になっています。

--allday-format-add-time, --allday-format-am12
終日スケジュール(時刻なし)のCSV出力形式を「0:00開始翌日0:00終了」とす
る。

オプション「--allday-format-add-time」指定時の終日スケジュール(時刻な
し)の出力
  "2025/09/05","00:00:00","2025/09/06","00:00:00" (1日間の終日スケジュール)
  "2025/09/05","00:00:00","2025/09/07","00:00:00" (2日間の終日スケジュール)

--allday-format-today-remove-time
終日スケジュール(時刻あり/時刻なし)の双方の出力形式を「時刻なしの、当
日開始、当日終了」とする。

※FlotingTimeのICSをOutlookにインポートすると、終日スケジュール(時刻なし)
が終日スケジュール(時刻あり)に変化します。本問題の差分をなくすため実装。

--allday-format-nexday-remove-time
終日スケジュール(時刻あり/時刻なし)の双方の出力形式を「時刻なしの、当
日開始、翌日終了」とする。

* 日付/時刻の表記関係

--day-format-slash-ymd
日付を日本式のスラッシュ区切りの「年/月/日」にします。
CSVのGaroonおよびOutlookClassic形式の場合のデフォルトです。

  例: 2025/12/31, 12:31

--day-format-iso8601-basic
日付をISO8601形式(基本)に変更します。

  例: 20251231, 1231

--day-format-iso8601-extended
日付をISO8601形式(拡張)に変更します。CSVのSimple形式の場合のデフォルトです。

  例 : 2025-12-31, 12:31

* メモ欄(DESCRIPTION)関係:

--delete-4th-line-onward
ICSメモ欄(description)の4行目以降を消してCSVに出力する。defaultでは消さない。

※詳細は関数ModCSV.modify_description()をみよ。

--show-teams-infomation
defaultではICSメモ欄(DESCRIPTION)のTeamsの会議インフォメーションを消
してCSVに出力する。パスワードが入ってるため。

本オプションを指定すると、本機能が無効になり、Teamsの会議インフォメー
ションを消さない。

※詳細は関数ModCSV.modify_description()をみよ。

--remove-tail-cr
タイトル(SUMMARY)やメモ欄(DESCRIPTION)の最後の改行や空白を除去する。
defaultでは除去しない。

※詳細は関数ModCSV.modify_description()をみよ。

* 繰返しスケジュールの一部上書き(RECURRENCE_ID)関係:

--show-hidden-schedules:
繰返しスケジュール(RRULE)の一部上書き(RECURRENCE_ID)を行った場合であっ
ても、ICS上は上書きされる前の情報が残っている場合がある。(カレンダーア
プリによって異なる。)

例えば、月から金の昼休みにスケジュールとして「昼食(社食)」を記載したあ
とに、水曜日を修正し「昼食(外食)」とした場合、内部的には水曜日のスケ
ジュールは「昼食(社食)」が残っている。

「--show-hidden-schedules」を指定すると一部修正する前のスケジュールも
表示する。

ただし上書きを行ったスケジュールと区別するため、Summaryにprefixとして
「Hidden: 」が挿入される。

具体的には水曜は「昼食(社食)」と「Hidden: 昼食(社食)」となる。

※詳細はRFC5545のRECURRENCE_ID参照。

--disable-recurrence-id
RFC5545のRECURRENCE_ID関連の処理を一切おこなわない。

* デバグ用:

--disable-exdate-format-bugfix
一部ICSファイルのRRULEのEXDATEの書式の問題でICSファイルの読み込みに失
敗する場合がある。EXDATEに修正を行う対策を行っている。

本対策を無効にする。

※詳細は関数PreSetup.bugfix_exdate_format()および misc/TECH-MEMO.txt をみよ。

--disable-naive-aware-mixed-bugfix
一部ICSファイルではDTSTART/DTENDが日付情報のみだが、UNTILにTimeZoneと
時間情報がある場合がある。本来はUNTIL側を修正すべきだが、本ソフトでは
DTSTART/DTENDに時刻情報を付与する修正を行う対策を行っている。

本対策を無効にする。

※pythonではTimeZone情報が無い時刻をnaive、有る時刻をawareと呼ぶ。

--DEBUG-UID="UID"
デバグ用。特定のUIDのオブジェクトを各種箇所で表示する。

* 煩雑なファイル確認:

--enable-file-exist-test
煩雑なファイル確認を有効にします。出力に指定されたCSVファイルがすでに
存在する場合は確認を求めます。入力に指定されたICSファイルがの日付が古
いと警告を出したり処理を停止します。

--disable-file-exist-test
-W
引数「--enable-file-exist-test」を無効にします。
"""

############################################
#設計がクソですが、Fは FeatureFlags型で、下記の外部公開の関数から呼び出された時に
#初期化します。関数から抜ける時に再びNoneに上書きします。
F = None
#

def parse_args(argv: list, amari_argv: int, allow_short_opt: str = None, \
               allow_long_opt: list = None) -> list:
    """
    引数の解析を行います。

    argv(list): コマンドで渡された引数

    amari_argv(int)はマイナスがついた引数解析後に残るマイナスがない引数の数です。


    多数オプションがありますので、制限されたコマンドのために
    allow_short_opt(str)およびallow_long_opt(list)を使います。与えられた
    引数のみ有効にします。

    返り値:
    オプション解析に失敗するとNoneを返します。

    正常終了時は引数解析後に残るマイナスがない引数と,FeatureFlags型のクラスです。

    """
    global F
    F = FeatureFlags()
    ret = PreSetup.parse_args(argv, amari_argv, allow_short_opt, allow_long_opt)
    flag = F
    F = None
    return ret, flag


def ics2csv(flag: FeatureFlags, ics_file_path: str, csv_file_path: str, timerange: int = 0) -> None:
    """
        ICS(iCalendar)ファイルをCSVファイルに変換する。

        引数:
            flag(FeatureFlags) 各種フラグ
            ics_file_path (str): 変換元のICS(iCalendar)ファイル。"stdin"を指定すると標準入力。
            csv_file_path (str): 変換先のCSVファイル。"stdout"を指定すると標準出力。
            timerange (int): CSVに変換する日時を限定する場合は、指定する。
                             2025年8月分がほしい場合は「202508」と指定する。
                             未指定や「0」だと全部変換する。
        返り値:
            None。失敗したら停止する。
    """
    global F
    F = flag
    ret = Main.ics2csv(ics_file_path, csv_file_path, timerange)
    F = None
    return ret

def guess_timerange(TIMERANGE: str, INPUT_ICS_FILENAME: str, OUTPUT_CSV_FILENAME: str) -> int:
    """
        ICSやCSVのファイル名よりCSVが出力する期間の値を推測します。

        引数:
        TIMERANGE:str
        INPUT_ICS_FILENAME:str
        OUTPUT_CSV_FILENAME:str

        返り値:
        CSV出力期間。失敗した場合はNoneを返します。
    """
    return TimeRange.guess(TIMERANGE, INPUT_ICS_FILENAME, OUTPUT_CSV_FILENAME)

############################################
__all__ = ('parse_args', 'ics2csv', 'guess_timerange',\
           'VERSION', 'HELP_LICENSE', 'HELP_PART1',\
           'HELP_PART2', 'HAIFU_URL', 'GITHUB_URL')

__v = tuple(sys.version_info)

# ライブラリvobjectはPython3.8以上が必要。
# (2025/10現在)Python3.9はすでにEOLのため本来なら3.10以上にしたいが
# macOSの標準Pythonが3.9なので3.9以上としている。
#
if __v < (3, 9):
    print(f"ERROR: Pythonはversion3.9以上が必要です。現在version {__v[0]}.{__v[1]}'", file=sys.stderr)
    sys.exit(1)
###

if __name__ == '__main__':
    F = FeatureFlags()
    help("libicsconvcsv")
#EOF
