# ICS to CSV コンバータ icsconvcsv インストール手順書

by 松元隆二

初版公開: 2025-9-24

最終更新:2026-1-22

**テキストエディタで閲覧される方へ**:本ファイルはMAKRDOWN言語で記述さ
れてます。9割プレインテキストですのでテキストエディタで閲覧頂いても問
題ありませんが、一点だけ補足すると「＜BR＞」は改行の命令になりますので
コマンド入力例などをコピペするときは除外ください。


はじめに README.md を読んでください。

# ソフトウエアのダウンロード

以下から最新版をzipファイルもしくはtar.gzファイルでダウンロードして
展開してください。

- <https://github.com/githubmatumoto/icsconvcsv/releases/>

ソフトウエアを展開したフォルダを覚えておいてください。ファイル
「INSTALL.txt」や「iscconvcsv.py」が含まれるフォルダです。

# 確認環境

## Ubuntu24.04.3LTS/日本語環境

  OS付属Pythonで確認しています。

> $ python3 --version <BR>
> Python 3.10.12

  OS設定によりPythonのバージョンが異なる可能性ありますがPython 3.9以上
  なら問題ないかと思います。

## Windows11Pro /日本語環境

  MicrosoftStore版Python3.13を導入して確認しています。

- <https://apps.microsoft.com/detail/9pnrbtzxmb4z?hl=ja-JP&gl=JP>

  ※URLは2025年12月現在。

コマンドプロンプトで下記コマンドを実行すると、導入されているバージョン
が確認できます。

> $ python3 --version <BR>
> Python 3.13.9

## macOS 26.0.1/ARM/日本語環境
  (IntelMacは未対応。動作確認機材無いため)

  Pythonが入ってない場合は「ターミナル」で「python3」と入力すると、OS
  標準のPythonの入手方法が表示されますので、それにしたがってください。

> % python3 --version <BR>
> Python 3.9.6

## (非対応)AlmaLinux8.10/日本語環境
  OS付属Pythonは3.6なので動作しません。

> $ python3 --version <BR>
> Python 3.6.8  (OS標準Pythonは非対応)

但し、OSの特権ユーザ等に相談頂いてPython3.9以降を導入すれば恐らく動
作すると思います。下記の手順になります。

> $ sudo dnf install python3.11

> $ python3.11 --version <BR>
> Python 3.11.13 (本ソフトウエアが対応するPython)

# 依存ライブラリのインストール。

ライブラリ vobjectを使わせていただいてます。バージョンは0.99で動作確認
しています。本ライブラリはPython3.8以降対応になります。

- <https://vobject.readthedocs.io/latest/>
- <https://py-vobject.github.io/>
- <https://github.com/py-vobject/vobject/releases>

## Windows(MicrosoftStore版Python3.13)

コマンドプロンプトでpython3/pip3がバージョン3.13以上であることを確認く
ださい。

> $ python3 --version <BR>
> $ Python 3.13.9

> $ pip3 --version <BR>
> pip 25.2 from C:\Program Files\WindowsApps\(省略)\pip (python 3.13)

ライブラリのインストール。

> $ pip3 install vobject <BR>
> (メッセージ省略)

成功したら下記のような表示になります。

> Successfully installed python-dateutil-2.9.0.post0 pytz-2025.2 six-1.17.0 vobject-0.9.9

(オプション) 通常は不要ですが、TimeZone関係のコードを有効にするには、
下記ライブラリを導入してください。

> $ pip3 install tzdata <BR>
> (メッセージ省略)

## Linux/macOSの設定

Linux/macOSではvenvの設定がお勧めです。前述してるがpythonのバージョン
は3.9以上である必要があります。

venv初期設定
  コマンド名「python3」はOSによって異なります

> $ cd ~/ <BR>
> $ python3 -m venv .icsconvcsv --prompt icsconvcsv

venv有効化

> $ source ~/.icsconvcsv/bin/activate

ライブラリのインストール

> $ pip3 install vobject <BR>
> (メッセージ省略)

成功したら下記のような表示になります。

> Successfully installed python-dateutil-2.9.0.post0 pytz-2025.2 six-1.17.0 vobject-0.9.9

※もしvobjectのバージョンが0.9.9より大きい数字になってた場合は、正常に
動作しない可能性があります。開発者まで連絡頂けると幸いです。

# icsconvcsvの動作確認

ソフトウエアを展開したフォルダにcdで移動してください。

> $ cd "ソフトウエアを展開したフォルダ"

以下のコマンドを実行してください。引数「-h」はヘルプメッセージを出力します。

> $ python3 icsconvcsv.py -h

正常に動作すると、下記のようなヘルプメッセージが出力されます。

```:text
Help on module icsconvcsv:

NAME
    icsconvcsv - ICS(iCalendar)をCSVに変換。

DESCRIPTION
    使用方法:

       $ python3 icsconvcsv.py [OPTION] 期間 入力.ics 出力.csv

       期間を指定して出力する例:
       $ python3 icsconvcsv.py 202512 calendar.ics schedules202512.csv
       $ python3 icsconvcsv.py guess calendar.ics schedules202509.csv
       $ python3 icsconvcsv.py all calendar.ics schedules-all.csv

       文字コードをShift_JISにする例:
       $ python3 icsconvcsv.py -Cshift_jis all calendar.ics schedules-sjis.csv

(以下省略)
```

ヘルプの停止は「q」を押してください。

                                                              以上です。
