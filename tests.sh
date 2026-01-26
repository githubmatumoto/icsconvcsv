#! /bin/bash
#
# Copyright (c) 2026 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
# コマンド「icsconvcsv.py」に様々なICSおよびオプションを与えて期待した
# 値が出力されるかのテストスクリプト。どのようなICSデータを与えてるか
# はREADME.test.md 参照ください。
#
# 補足: スクリプトで nkf を使ってます。nkfをインストールください。

PYTHON=python3
PROGNAME=./icsconvcsv.py

TIMERANG=all

TMP1CSV=./tmp1.csv
TMP2CSV=./tmp2.csv
TMPLOG=./log.txt

# == 0 エラー時に停止
# == stop エラー時に続行
#ERROR_TAIOU=stop
ERROR_TAIOU=continue
# off: nkfなし
# on: nkfあり
NKF=off

# off: CSV 正規化なし
# on: CSV 正規化あり
NORMAL=off

PROG_NORMAL=./normal_csv.py
# 上記プログラムで表示する行数
# -1, -2, -3, -4, -5,
# 無指定もしくは-aなら全部
PRINT_LINE=

#on: 煩雑なログをだす。
#off: エラー時のみログをだす。
SILENT=on

function cmp_ics() {
    ARGS=$1
    ICS=ICS/$2."ics"
    if [ $# -eq 3 ]; then
	CSV=CSV/$3."csv"
    elif [ $# -eq 2 ]; then
	CSV=CSV/$2."csv"
    else
	echo "引数エラー"
	exit
    fi

    if [ ! -f ${ICS} ]; then
	echo "ファイル" ${ICS} "が存在しません。"
	exit
    fi

    if [ ! -f ${CSV} ]; then
	echo "ファイル" ${CSV} "が存在しません。"
	exit
    fi

    if [ -f ${TMP1CSV} ]; then
       rm ${TMP1CSV}
    fi

    if [ -f ${TMP2CSV} ]; then
       rm ${TMP2CSV}
    fi

    if [ -f ${TMPLOG} ]; then
       rm ${TMPLOG}
    fi

    if [ $SILENT == "off" ]; then
	echo -n "CHECK: > ${PYTHON} ${PROGNAME} ${ARGS} ${ICS} ${TMP1CSV}"
    fi
    ${PYTHON} ${PROGNAME} ${ARGS} ${ICS} ${TMP1CSV} 2> ${TMPLOG}
    retval=$?

    if [ $retval -ne 0 ] ; then
	if [ $SILENT == "on" ]; then
	    echo -n "CHECK: > ${PYTHON} ${PROGNAME} ${ARGS} ${ICS} ${TMP1CSV}"
	fi

	echo 'ERROR: 失敗しました(終了ステータス異常)。'
	echo "-- ERROR LOG --------------------------"
	cat -n ${TMPLOG} | tail
	echo "---------------------------------------"
	if [ $ERROR_TAIOU = "stop" ]; then
	   exit
	fi
	return
    fi

    if [ ! -f ${TMP1CSV}  ] ; then
	echo 'ERROR: 失敗しました(ファイル生成なし)。'
	if [ $ERROR_TAIOU = "stop" ]; then
	   exit
	fi
	return
    fi


    if [ -f ${TMPLOG} ]; then
       rm ${TMPLOG}
    fi

    if [ $NKF = "on" ];then
	# 自動判断はさせず、入力SJIS、出力UTF-8固定。
	#echo "CONV: > nkf -S -w < ${TMP1CSV} > ${TMP2CSV}"
	#echo "CHECK: USE NKF"
	nkf -S -w < ${TMP1CSV} > ${TMP2CSV}
	rm ${TMP1CSV}
	mv ${TMP2CSV} ${TMP1CSV}
    fi

    #echo "CHECK: > diff -u ${TMP1CSV} ${CSV}"

    if [ $NORMAL = "on" ]; then
	#echo "CHECK: USE NORMALIZE_CSV"
	${PYTHON} ${PROG_NORMAL} ${PRINT_LINE} < ${TMP1CSV} > ${TMP2CSV}
	rm ${TMP1CSV}
	mv ${TMP2CSV} ${TMP1CSV}
	${PYTHON} ${PROG_NORMAL} ${PRINT_LINE} < ${CSV} > ${TMP2CSV}

	diff -u ${TMP1CSV} ${TMP2CSV} > ${TMPLOG}
	retval=$?
    else
	diff -u ${TMP1CSV} ${CSV} > ${TMPLOG}
	retval=$?
    fi

    if [ $retval -ne 0 ] ; then
	echo "CHECK: > diff -u ${TMP1CSV} ${CSV}"
	echo 'ERROR: 失敗しました(差分あり。差分は冒頭10行のみ)'
	echo "-- ERROR LOG --------------------------"
	cat -n ${TMPLOG} | head
	echo "---------------------------------------"

	if [ $ERROR_TAIOU = "stop" ]; then
	   exit
	fi
	return
    fi

    if [ $SILENT == "off" ]; then
	echo ": SUCCESS "
    fi
}

echo "ライブラリicsconvcsvの一括テストスクリプト。「MEMO:失敗で正常」とある場合は無視して問題ありません。"

if [ $SILENT == "off" ]; then
    echo "テストに失敗時のみ詳細がでます。"
fi

ERROR_TAIOU=stop
echo
echo "MEMO: EXDATEの書式の問題で例外を送出するバグフイックス"
cmp_ics "all" "ga1"

ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: EXDATEの書式の問題で例外を送出するバグフイックスを無効化。"
cmp_ics "--disable-exdate-format-bugfix all" "ga1"
ERROR_TAIOU=stop

echo
echo "MEMO: 日付情報でnaiveとawareが混在するバグフイックス"
cmp_ics "all" "ou1"

ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: 日付情報でnaiveとawareが混在するバグフイックスを無効化。"
cmp_ics "--disable-naive-aware-mixed-bugfix all" "ou1"
ERROR_TAIOU=stop

echo
echo "MEMO: Garoonが生成するICSファイルで、ESCが適切に行われない例。"
echo "MEMO: 記号(,と;)の適切なエスケープが行わない。Outlookは成功。"
cmp_ics "-FGaroon -Cutf-8 all" "ou2"
cmp_ics "-FGaroon -Cutf-8 all" "ouc2" "ou2"
ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: garoonのみ失敗する"
cmp_ics "-FGaroon -Cutf-8 all" "ga2" "ou2"
ERROR_TAIOU=stop

echo
echo "MEMO: TimeZoneを誤指定した時の挙動。"
cmp_ics "all" "ou3"

ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: TimeZoneを誤指定。6月2日が欠落。"
cmp_ics "-TUS/Eastern all" "ou3"
echo
echo "MEMO: 失敗で正常: TimeZoneを誤指定。7月20日の除外を失敗。7月27日が欠落。"
cmp_ics "-TUS/Eastern all" "ou1"
ERROR_TAIOU=stop

echo
echo "MEMO:Outlookで使える属性の調査。出力はCSV-Simple形式のみ調査。"
cmp_ics "all" "ou10"
cmp_ics "all" "ouc10"
cmp_ics "all" "ou10-free"
cmp_ics "all" "ou10-us" "ou10-free"

echo
echo "MEMO:いろんな例のテスト ou11.icsの出力と比較"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ou11"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ouc11" "ou11"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ga11" "ou11"

echo
echo "MEMO: Garoonが生成したCSV(ga11-org)と比較"
echo "MEMO: 失敗で正常: TEST:08のみ差分がでる。outlookでは入力できない時刻指定"
ERROR_TAIOU=continue
NORMAL=on
cmp_ics "-Fgaroon -Cutf-8 -m all" "ga11" "ga11-org"
NORMAL=off
ERROR_TAIOU=stop

echo
echo "MEMO:拡張業務番号"
cmp_ics "-Fgaroon -Cutf-8 -m -z all" "ou12"
cmp_ics "-Fgaroon -Cutf-8 -m -z all" "ouc12" "ou12"

NORMAL=on
PRINT_LINE=-1
cmp_ics "-Fgaroon -Cutf-8 -m -z all" "ou12-limit2" "ou12"
cmp_ics "-Fgaroon -Cutf-8 -m -z all" "ouc12-limit2" "ou12"
NORMAL=off
PRINT_LINE=

echo
echo "MEMO:タイトル(SUMMARY)の分割"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ou13"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ouc13" "ou13"

ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: TEST:84のみ差分がでる。Garoonの記号(,と;)の適切なエスケープが行わないのバグ"
cmp_ics "-Fgaroon -Cutf-8 -m all" "ga13" "ou13"
ERROR_TAIOU=stop

echo
echo "MEMO:タイトル(SUMMARY)の分割停止"
cmp_ics "-Fgaroon -Cutf-8 --disable-split-summary all" "ou13" "ou13-dis-sum"
cmp_ics "-Fgaroon -Cutf-8 --disable-split-summary all" "ouc13" "ou13-dis-sum"

ERROR_TAIOU=continue
echo
echo "MEMO: 失敗で正常: TEST:84のみ差分がでる。Garoonの記号(,と;)の適切なエスケープが行わないのバグ"
cmp_ics "-Fgaroon -Cutf-8 --disable-split-summary all" "ga13" "ou13-dis-sum"
ERROR_TAIOU=stop

echo
echo "MEMO: 時刻の表示の確認(day-fomart)"

function ics_day_format() {
    opt=$1
    suf=$2
    opt="${opt} -Fgaroon -Cutf-8"
    cmp_ics "${opt} --day-format-slash-ymd all" "ou14" "ou14-slash-${suf}"
    cmp_ics "${opt} --day-format-iso8601-basic all" "ou14" "ou14-basic-${suf}"
    cmp_ics "${opt} --day-format-iso8601-extended all" "ou14" "ou14-extended-${suf}"

    cmp_ics "${opt} --day-format-slash-ymd all" "ouc14" "ou14-slash-${suf}"
    cmp_ics "${opt} --day-format-iso8601-basic all" "ouc14" "ou14-basic-${suf}"
    cmp_ics "${opt} --day-format-iso8601-extended all" "ouc14" "ou14-extended-${suf}"

    cmp_ics "${opt} --day-format-slash-ymd all" "ga14" "ou14-slash-${suf}"
    cmp_ics "${opt} --day-format-iso8601-basic all" "ga14" "ou14-basic-${suf}"
    cmp_ics "${opt} --day-format-iso8601-extended all" "ga14" "ou14-extended-${suf}"

    cmp_ics "${opt} --day-format-slash-ymd all" "ou14-us" "ou14-slash-${suf}"
    cmp_ics "${opt} --day-format-iso8601-basic all" "ou14-us" "ou14-basic-${suf}"
    cmp_ics "${opt} --day-format-iso8601-extended all" "ou14-us" "ou14-extended-${suf}"
}

ics_day_format "--allday-format-today" "today"
ics_day_format "--allday-format-nextday" "nextday"
ics_day_format "--allday-format-am12" "am12"
ics_day_format "--allday-format-today-remove-time" "today-rt"
ics_day_format "--allday-format-nextday-remove-time" "nextday-rt"

echo
echo "MEMO: 時刻の表示の確認(TimeZoneあり)"

cmp_ics "--show-timezone -Fgaroon -Cutf-8 all" "ou14" "ou14-tz"
#GaroonのICSはTimeZoneがないのでなしで正常
cmp_ics "--show-timezone -Fgaroon -Cutf-8 all" "ga14" "ou14-slash-today"
cmp_ics "--show-timezone -Fgaroon -Cutf-8 all" "ouc14" "ou14-tz"
cmp_ics "--show-timezone -Fgaroon -Cutf-8 all" "ou14-us" "ou14-us-tz"

echo
echo "MEMO: CSV書式の確認"

cmp_ics "--print-csv-header all" "ou14" "ou14-sim"
cmp_ics "--print-csv-header all" "ga14" "ga14-sim"
cmp_ics "--print-csv-header all" "ouc14" "ou14-sim"
cmp_ics "--print-csv-header all" "ou14-us" "ou14-sim"

cmp_ics "--print-csv-header -Fgaroon -Cutf-8 all" "ou14" "ou14-ga"
cmp_ics "--print-csv-header -Fgaroon -Cutf-8 all" "ga14" "ou14-ga"
cmp_ics "--print-csv-header -Fgaroon -Cutf-8 all" "ouc14" "ou14-ga"
cmp_ics "--print-csv-header -Fgaroon -Cutf-8 all" "ou14-us" "ou14-ga"

cmp_ics "--print-csv-header -Foutlookclassic all" "ou14" "ou14-ouc"
cmp_ics "--print-csv-header -Foutlookclassic all" "ga14" "ga14-ouc"
cmp_ics "--print-csv-header -Foutlookclassic all" "ouc14" "ou14-ouc"
cmp_ics "--print-csv-header -Foutlookclassic all" "ou14-us" "ou14-ouc"

echo
echo "MEMO: アメリカ東海岸(EDT)の時刻の確認"

cmp_ics "--show-timezone -Fgaroon -Cutf-8 all" "ou15-us" "ou15-us"
cmp_ics "--show-timezone -TAsia/Tokyo -Fgaroon -Cutf-8 all" "ou15-us" "ou15-jp"


echo
echo "MEMO: Teams会議およびRECURRENCE_ID命令関連"
cmp_ics "-Fgaroon -Cutf-8 all" "ou16" "ou16"
# ホントはICS側に空白一つ分の差分あったが、ICS側を修正しています。
cmp_ics "-Fgaroon -Cutf-8 all" "ouc16" "ou16"

cmp_ics "--show-hidden-schedules -Fgaroon -Cutf-8 all" "ou16" "ou16-hidden"
cmp_ics "--show-hidden-schedules -Fgaroon -Cutf-8 all" "ouc16" "ou16-hidden"

cmp_ics "--disable-recurrence-id -Fgaroon -Cutf-8 all" "ou16" "ou16-dis-rec"
cmp_ics "--disable-recurrence-id -Fgaroon -Cutf-8 all" "ouc16" "ouc16-dis-rec"

cmp_ics "--show-teams-infomation -Fgaroon -Cutf-8 all" "ou16" "ou16-teams"
cmp_ics "--show-teams-infomation -Fgaroon -Cutf-8 all" "ouc16" "ouc16-teams"

cmp_ics "--show-teams-infomation --delete-4th-line-onwar  -Fgaroon -Cutf-8 all" "ou16" "ou16-4th"
cmp_ics "--show-teams-infomation --delete-4th-line-onwar -Fgaroon -Cutf-8 all" "ouc16" "ou16-4th"

echo
echo "MEMO: 文字コード変換テスト(ICSファイル側にShift_JISに変換できない文字があると差分となる)"

NKF=on
cmp_ics "--show-teams-infomation -Fgaroon all" "ou16" "ou16-teams"
cmp_ics "--show-teams-infomation -Fgaroon all" "ouc16" "ouc16-teams"
NKF=off

echo
echo "MEMO: RDATE関係"
# 作業メモ「make gen-ouc-omitdes.csv」の出力がほぼ同等のはず。
cmp_ics "-Fomitdescription all" "ouc17-limit2"


echo
echo "正常終了しました。"


if [ -f ${TMP1CSV} ]; then
   rm ${TMP1CSV}
fi
if [ -f ${TMP2CSV} ]; then
   rm ${TMP2CSV}
fi
if [ -f ${TMPLOG} ]; then
   rm ${TMPLOG}
fi
