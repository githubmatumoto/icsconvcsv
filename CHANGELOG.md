-*- text -*-
# Cangelog:


## 2025-9-24: Version:1.0 (公開終了)

初版公開

## 2025-9-25: Version:1.1 (公開終了)

- 用語統一
  変数名等でicalとicsが混ざってたので、icsに統一。

## 2025-9-26: Version:1.2 (公開終了)
(内部メモ:subversion revision 2070, フォルダv1.2)

- CSVの改行コードの扱いが悪く無駄な空行が入るのを改善。生成される改行
  コードについては misc/TECH-MEMO.txt を参照ください。

- CSVでダブルクオートで囲まれた文字列の最後に改行が空白があった場合除
  去するコードを追加。

- 出力文字コードがShift_JISの時にUTF-8からShift_JISに変換する作業に失敗
  した時のエスケープ手段がstdoutに出力した時とファイルに出力した時で異なっ
  たので、統一。

- macOS 13.7.8/Intel/日本語環境動作確認しました。

- ライブラリのバージョン管理用の変数を追加しています。

- その他用語の統一など行ってます。

## 2025-9-30: Version:1.3 (非公開)
(内部メモ:subversion revision 2079, フォルダv1.3)

- 上書スケジュール(RECURRENCE-ID)へ限定対応。細かいRECURRENCE-ID命令に
  は対応していない。挙動がおかしい場合はオプション「-w」で無効化された
  命令をCSVに出力するようにして欲しい。

- 上書スケジュール(RECURRENCE-ID命令)の処理を無効にするオプション「-x」
  を追加。

- 上書スケジュール対応のため、CSVに出力する前に一度バッファリングを行
うように変更。

- CSVに出力時の特殊な値を定義 <BR>
  ここで述べる値をカレンダーアプリに直接記載すると誤動作する。
  - 「(N/A)」: 指定の要素がVEVENTに無かった。SUMMARY/DESCRIPTONのみ。
  ICSアプリにより、上記要素が無い場合、下記の挙動がある。<BR>
> 要素を未定義とする場合: 「(N/A)」が入る。<BR>
> 改行をいれる場合: 空行が入る。
  - 「(REFERENCE DATA DOES NOT EXIST)」: 滅多にないはずだが、
  RECURRENCE-IDで上書をするVEVENTにはSUMMARY/DESCRIPTONが無い場合がある。
  基のVEVENTからコピーする処理を行っているが、基のVEVENTが見つからなかっ
  た。
  - 「Hidden: 」オプション「-w」を指定した時のみ。<BR>
  RECURRENCE-IDで上書をした基のVEVENTのSUMMARYに「Hidden: 」を付けて出力する。

- 暗にTimeZoneとしてJSTを想定しているコードに 「*DEPEND ON JST*」とい
  う印をつけた。(把握している分のみ)

## 2025-12-10: Version:1.4(非公開)
(内部メモ:subversion revision 2100, フォルダv1.4。間違えてv1.4版の
「libics2gacsv.py」を更新してしまったので2105で差し戻しを行っている。
なので、「libics2gacsv.py」は現在2105だが、2100と同じファイルのはず)

- 用語統一
   Windowsアプリ版の「Outlook(classic)」を「Outlook(legacy)」と誤記。
    →正しい「Outlook(classic)」に修正。

> 「日程」→「スケジュール」に統一。
> 日本語の「時間」と「時刻」の誤用→正しい日本語の意味に修正。

- 概ね下記のICSファイルに対応。
> Cybozu Garoon(Version 5.0.2) <BR>
> Web版 Outlook <BR>
> Windowsアプリ版 Outlook(classic) <BR>
> いずれも2025年12月ごろに生成されたICSファイルの出力で確認。

- 警告やメッセージはv1.3まではSTDOUTに出力してたが、STDERRに出力するよ
  うに修正。

- 上書スケジュール(RECURRENCE-ID)の対応強化。概ね確認できた範囲では正
  常に変換できる。

- 開始時刻が0:00、終了時刻が翌日0:00のスケジュールを終日スケジュールと
  みなすオプション「-g」を追加。

- (不完全)旧版では常にTimeZoneとしてJST(Asia/Tokyo)を想定していたが、
  他のTimeZoneに対応。真面目に調査していない。

  v1.3で記載した「*DEPEND ON JST*」削除。

- TimeZone情報はICSファイルのVTIMEZONEというエリアに記載されているが、
  VTIMEZONEが無いICSファイルは「Flating Time」と呼ばれる。
  その場合、以下の情報をSTDERRに出力する。

```:text
  "INFO: ICSデータにTimeZoneデータがありません。"
  "INFO: Floating Timeのデータです。(Ref: RFC5545, 3.3.12. TIME)"
```

- 一部ICSファイルはVTIMEZONEがにTimeZoneが2個以上定義されている。
  その場合、以下の情報をSTDERRに出力する。

```:text
   "INFO: ICSファイルにTimzeZoneが複数定義されています。"
   "INFO: 現在定義されているTimeZone一覧: {一覧表示}"
   "INFO: TimeZoneとして1番目に定義されている[初出のTimeZone]を採用します。"
   "WARNING: 採用したTimeZoneが不適切な場合、繰返しスケジュールの最終日(UNTIL)の計算に失敗し、"
   "WARNING: スケジュールが欠落する可能性あります。不適切な場合は引数で指定してください。"
```

   UNTILの時刻はGMTで記載されていることがあります。そのため、文脈によっ
   てはTimeZone情報が必須となります。

   複数定義されている場合、TimeZoneの指定を行うオプション「-T」を追加。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」加筆。

- macOS(Intel)の動作確認環境が無くなったため、サポートOSから削除
- macOS 26.0.1/ARM/日本語環境をサポートOSに追加。

- Pythonのサポートバージョンを3.9以降に修正。
  Python3.9はすでにEOLのため本来なら3.10以上にしたいが
  macOSの標準Pythonが3.9なので3.9以上としています。。

- サポートOSからAlmalinux8.10を削除しました。
  OS標準のPythonが3.6のため。ただしOS標準パッケージにPython3.11が含ま
  れるため、インストールすれば、動作可能。

## 2025-12-18: Version:2.0 (公開終了)
(内部メモ:subversion revision 2124, フォルダv2.0)

- 公開用にversion打ち直し。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」を削除して、Qiitaに移動。
  https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8

- Ubuntuのサポートバージョンを22.04LTSから24.04LTSに変更。
  上記は誤記修正になります。開発用に使っていた環境を22.04と誤認してい
  たため。リリース版公開に合わせて再確認を行ったところ24.04でした。

- コード内の各種コメントの精査。

	- 変更履歴をソースコード内から misc-CHANGELOG.md に移動。

## 2025-12-26: Version:2.1 (公開終了)
(内部メモ:subversion revision 2151, フォルダv2.1)

- 職場の業務記録提出用のスクリプト kiroku.py 追加。ファイル名を決
  め打ちで生成する。

- SUMMARYに業務番号を記載する、拡張追加。
```:text
  libics2gacsv.py:
  関数 modify_enhanced_gyoumunum()　追加
  制御変数 flag_enhanced_gyoumunum 追加

  ics2gacsv.py:
  引数 -z 追加(flag_enhanced_gyoumunum=True)
```

- 上書スケジュール(RECURRENCE-ID)のバグ対応。

Outlookが生成したICSファイルでもとの繰返し(RRULE)スケジュールのDESCRIPTIONが未定義"(N/A)"であり、
上書きスケジュールのDESCRIPTIONが未定義"(N/A)"であった場合は、
誤ってDESCRIPTIONに"(REFERENCE DATA DOES NOT EXIST)"が入いるようになった。

```:text
libics2gacsv.py:
def modify_reference_id_data()
   誤:     if csv_buffer[i][k] is None:
   正:     if (csv_buffer[i][k] is None) or (outlook_bugfix and csv_buffer[i][k] == UNREF):
```

- 修正split_garoon_style_summary
```:text
  変数名修正
  旧:flag_matumoto_modify

  新: flag_split_summary_enhance
  具体的な追加項目を今まで関数内にハードコーディングしてたが、ics2gacsv.pyで
  下記変数に代入する形に修正
  G_SPLIT_SUMMARY_ENHANCE
```

## 2026-1-26: Version:3.0beta1
(内部メモ:subversion revision 2193, フォルダv3.0)

- ソフトウエア名をics2gacsvからicsconvcsvに変更。 ソフトウエア名に商標
であるCybozu Garoonの一部である「ga」が含まれていたため。

- README.txtおよびINSTALL.txtをプレインテキストからMAKRDOWN言語に変更。
ファイル名をREADME.mdおよびINSTALL.mdに変更。9割プレインテキストですの
でテキストエディタで閲覧頂いても問題ありませんが、一点だけ補足すると
「＜BR＞」は改行の命令になりますのでコマンド入力例などをコピペするとき
は除外ください。

- Pythonのclassを導入しています。ただし名前空間を綺麗にするためだけに
使っていますので、オブジェクト指向は用いてません。

  また関数名の整理をおこなっています。そのため、CHANGELOG.mdに記載の
  v2.1以前の修正についても該当する関数が現存しない場合があります。

- 関数名で時刻情報のnaiveをnativeと誤記してたので修正。

- 多数の文章でICSをISCと誤記してたので修正。

- 引数解析関数を独立させ(PreSetup.parse_args())ライブラリ側に統合。引
  数を解説するヘルプの文章もライブラリ側に統合。

  ライブラリ呼び出し側のコマンド(icsconvcsv.py, ics2gacsv.py, kiroku.py)から
  PreSetup.parse_args()を呼び出す時に、デフォルト引数を設定することにより、
  コマンドの挙動が変わるような構成になってます。

- 文字コードを指定するオプション 「-C"文字列"」追加。
  -Cshift_jis, -Cutf_8, -Cutf_8_sig

  utf_8_sigはBOM(Byte Order Mark)を付与する。

  文字コードを指定するオプション 「-u 」削除。

- CSVの出力はGaroonとほぼ同等の出力のみ対応していたが、
  複数の出力形式に対応。出力形式は「-F"文字列"」で指定する。

- CSVの出力のデフォルトを独自定義の「Simple」に変更。
  Garoonとほぼ同等の出力に変更する場合は「-FGaroon」を
  指定する。

- デフォルトの文字コードをShift_JISからUTF-8に変更。但し「-FGaroon」を
指定するとデフォルトはShift_JIS。

- CSVの終日スケジュールの時刻形式を指定するオプションの整理

  廃止: -o, -g

  追加:　--allday-format-*

- CSVの日時のフォーマットのオプション追加。

  日本式のslashで区切るのに加えてISO8601形式に対応しています。

  --day-format-*

- 多くのショートオプション(1文字)廃止。ロングオプション(複数文字)に変更。

- CSVのヘッダ(1行めの項目)を表示するオプション追加。
  --print-csv-header, -k

- 出力ファイルの上書き確認/入力ファイルの日付確認を行なうオプション追加。
  --enable-file-exist-test

- テスト用に同梱してたdebug.shを廃止。代わりに複雑な動作確認を行う
  tests.shを同梱。説明についてはREADME.tests.md参照ください。

- RDATE命令対応。但しカレンダーソフトがRDATEを滅多に出力しないため、
  テストが不完全と考えています。

## 2026-1-27: Version:3.0
(内部メモ:subversion revision 2196, フォルダv3.0)

- 繰返し命令の処理のサンプル ouc4.ics 追加。

- 職場向け仕様である拡張業務番号を拡張登録番号に改名。職場の用語変更に
  伴う修正。ただし関数名やサンプルなどは修正していません。

  引数追加。--enhance-tourokunum, --enhance-touroku-number

- 引数--add-summary-head に使えない文字列「Hidden」を追加。

# Known bugs:

- 西暦を判断する基準の正規表現が「[^\\d]20[\\d]{6}」などになってるので、
西暦2000年から2099年までしか動作しない。

- 上書スケジュール(RECURRENCE-ID)で上書きされる元のスケジュールが1999
年以前もしくは2100年以降だと動作しない。

- (未調査)サマータイムの切り替えがあるTimeZoneの場合、サマータイムの前
後で1日の長さが23時間もしくは25時間の場合はおそらく正常に動作しない。

- Teamsの会議インフォーメーションの削除はフォーマットが変わったら無効。
2025年9月のフォーマットを元に削除を行います。

- ICSの各要素になにも入ってないというのを示すのに"(N/A)"という文字列を
使っています。SUMMARYやDESCRIPTIONに最初から"(N/A)"と入っていた場合は
誤動作する。

- ICSの上書スケジュール(RECURRENCE-ID)繰返し要素の処理に
"(REFERENCE DATA DOES NOT EXIST)"という文字列を使っています。SUMMARYや
DESCRIPTIONに最初から上記が入って場合は誤動作する。

- 登録番号を暗に4桁と想定している。5桁以上なら下記関数の正規表現を修正
する。

> ModCSV.enhanced_gyoumunum()

- RDATEに対応はしたが、動作確認例が少ないため、要注意。
