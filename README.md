# ICS to CSV コンバータ icsconvcsv

by 松元隆二

初版公開2025-9-24

最終更新:2026-1-22

**テキストエディタで閲覧される方へ**:本ファイルはMAKRDOWN言語で記述さ
れてます。9割プレインテキストですのでテキストエディタで閲覧頂いても問
題ありませんが、一点だけ補足すると「＜BR＞」は改行の命令になりますので
コマンド入力例などをコピペするときは除外ください。

# はじめに

スケジュールの標準フォーマットであるICS形式(別名iCalendar形式)をCSV形
式に変換するツール「icsconvcsv」です。

旧版v2.*では「ics2gacsv」という名前で公開してました。

配布元は以下になります。

- <https://qiita.com/qiitamatumoto/items/ab9e0cb9a6da257597a4>

- <https://github.com/githubmatumoto/icsconvcsv>

ICS形式については下記を参照ください。

<https://ja.wikipedia.org/wiki/ICalendar>

入力のICSファイルは、以下のカレンダーアプリが生成するICSファイルに対応
しています。但し主にWeb版Outlookを主眼に開発を行ってます。

- Cybozu Garoon(Version 5.0.2)
- Web版 Outlook
- Windowsアプリ版 Outlook(classic)

いずれも2026年1月ごろに生成されたICSファイルの出力で確認。

出力するCSVの形式は、

- 独自形式(simple形式)
  - CSV項目: "DTSTART:DAY", "DTSTART:TIME", "DTEND:DAY" "DTEND:TIME",
              "SUMMARY" , "DESCRIPTION", "X-MICROSOFT-CDO-BUSYSTATUS",
	      "CATEGORIES"
  - 文字コード:UTF-8
  - 概ねICSの要素のベタ展開。
  - 将来CSV項目を後ろに追加する形で増やす可能性あります。

- Cybozu Garoon形式
  - CSV項目: "開始日", "開始時刻", "終了日", "終了時刻", "予定",
    "予定詳細", "メモ"
  - 文字コード:Shift-JIS

に対応してます。

## 開発動機

職場ではグループウエアCybozu Garoonのスケジュールが生成するCSV形式で業
務記録簿を提出することになってます。グループウエアがOffice365に変更に
なる事になりました。Office365に含まれるOutlook(classic)CSV形式の出力に
対応していますが、利用者が少なく、アプリのインストールが不要である
Outlook(Web版)ではスケジュール表をエキスポートする場合はICS形式しか選
べないため、CSV形式に変換する必要がありました。

内部向けツールですが、ICS to CSVを行いたい需要があるかもしれないので、
公開します。

資料内で「業務記録」という単語が多数出てきますが、同じ職場以外の皆様は、
本ツールで作成したCSV形式のファイルとお考えください。

なお、まだ実運用は行ってない模様。。これからバグ出しですね(^^;

上記のような目的で作成されたプログラムのため、Outlook(web版)で直接CSV
の出力が可能になった場合、開発は終了する可能性があります。

#依存ライブラリ
ライブラリ vobjectを使わせていただいてます。バージョンは0.99で動作確認
しています。

- <https://vobject.readthedocs.io/latest/>

- <https://py-vobject.github.io/>

- <https://github.com/py-vobject/vobject/releases>

*ライセンス

ライセンスは Apache License 2.0です。

依存しているライブラリvobjectと同様のライセンスになります。

*動作環境およびインストール方法

以下から最新版をzipファイルもしくはtar.gzファイルでダウンロードして
展開してください。

- <https://github.com/githubmatumoto/icsconvcsv/releases/>

展開したファイルに含まれる

- INSTALL.md (テキストエディタで閲覧ください)

を参照して、Pythonのライブラリのインストールを行ってください。

ソフトウエアを展開したフォルダを覚えておいてください。ファイル
「INSTALL.md」や「iscconvcsv.py」が含まれるフォルダです。

# 使い方

## スケジュールの記入とICSファイルのダウンロード

Outlook(web版)でいろいろスケジュールを作成してください。そのあと、ICS
形式でダウンロードしてファイルに保存してください。ICSのダウンロード手
順は下記の資料を参考にしてください。

- <https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8>

を参照ください。

## 毎回必要な初期設定や確認事項

本プログラムを利用する前に、``毎回''必要な初期設定や確認事項があります。

Linux/macOSは以下を実行してください。Pythonの初期化になります。


> $ source ~/.icsconvcsv/bin/activate


Windowsはライブラリvobjectを導入したpythonと同じであるか確認ください。

> $ python3 --version

## ICSをCSVに変換

ICSをCSVに変換するコマンドは下記の形式になります。

> $ python3 icsconvcsv.py 期間 入力ICSファイル 出力CSVファイル

### 実行例1:

ソフトウエアを展開したフォルダにICSファイルをcalendar.icsという名前
で置いてください。

CSVに出力する「期間」は年4桁と数字2桁で指定してください。下記の実行例
では202509と指定していますが、2025年9月分を指定した事になります。

下記の実行例ではICSファイル「calendar.ics」に含まれる「2025年9月分」の
スケジュールをCSV形式に変換しCSVファイル「schedules202509.csv」に出力
します。

> $ cd "ソフトウエアを展開したフォルダ"

> $ python3 icsconvcsv.py 202509 calendar.ics schedules202509.csv

変換に成功した場合

> INFO: 変換に成功しました: 'calendar.ics' to 'schedules202509.csv'

と表示されます。

CSVファイルの文字コードはUTF-8になっています。Excelで閲覧するとおそら
く文字化けしています。Shift-JISにする場合は引数 「-Cshift_jis」を追加
してください。

### 実行例2:
作者の職場の業務記録提出用のスクリプトです。入力/出力ファイル名が決め
打ちになっています。

ソフトウエアを展開したフォルダにICSファイルをcalendar.icsという名前で
置いてください。以下のコマンドで提出用のCSV業務記録簿が2個生成されます。
今月分と先月分が生成されます。

> $ cd "ソフトウエアを展開したフォルダ"

> $ python3 kiroku.py "業務記録簿提出者名"

利用方法の詳細は引数「-h」を渡して実行してみてください。

> $ python3 kiroku.py -h

※作者の職場の業務記録提出用のため頻繁に仕様が変更になります。

### 旧版の互換コマンド

過去に公開してた「ics2gacvs.py」の互換コマンドを同梱しています。

利用方法の詳細は引数「-h」を渡して実行してみてください。

> $ python3 ics2gacsv.py -h

大部分のオプション引数は廃止してます。細かい指定を行う場合は、
「icsconvcsv.py 」をお使いください。

### サンプルスクリプト:

コマンドの実行例をいくつか記載したサンプルスクリプトを以下の名前で作成
してます。ICSファイルをcalendar.icsという名前で置いて実行してみてくだ
さい。※2025年12月と2026年1月のスケジュールが無い場合は空のファイルが
生成されます。

Windows用 : Windowsバッチファイル

> $ samplegen.bat

ShiftJISの下記のファイルが生成されます。

- sjis-202512.csv : 2025年12月のスケジュール (CSVのsimple形式)
- sjis-202601.csv : 2026年1月のスケジュール
- sjis-all.csv    : すべてのスケジュール
- garoon-202512.csv : 2025年12月のスケジュール (CSVのGaroon形式)
- garoon-202601.csv : 2026年1月のスケジュール
- garoon-all.csv    : すべてのスケジュール

Linux,macOS用: Shellスクリプト

> $ sh samplegen.sh

上記のWindowsバッチファイルが生成するファイルに加えて以下のファイルが生成されます。

- utf8-202512.csv : 2025年12月のスケジュール (CSVのsimple形式)
- utf8-202601.csv : 2026年1月のスケジュール
- utf8-all.csv    : すべてのスケジュール

Excelで閲覧や編集する場合は文字コードがShiftJISのファイル
(sjis-*.csv)を使ってください。

※セキュリティエラーが出たら中身を1行づつコピペして実行してください。

## ICSとCSVの要素の対応について

ICSとCSVの要素の対応については下記blogを参照ください。

- <https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8>

## Known bugs

既知のバグがいろいろあります。CHANGELOG.md 参照ください。

特に、業務で使う場合は充分な事前テストをお願いします。ほぼ同等のスケ
ジュールをWeb版OutlookとCybozu Garoonに記載し、比較調査お願いします。

またCybozu GaroonにはCSV以外にICSを出力する機能があります。GaroonでICS
を出力し、icsconvcsvでCSVに変換して、比較していただくのも良いかもしれま
せん。

特に

- 終日スケジュール
- 繰り返しスケジュールのTimeZone
- 繰り返しスケジュールの一部上書(RECURRENCE-ID)

でいろいろバグが出ました。一応既知の例はデバックしましたが、未知のバグ
があるかもしれません。

以上です。
