# Backlogcsv2xlsx

## 概要
- Backlogからダウンロードできるcsvデータをもとにクロスミーティング用資料のテーブル(xlsx形式)を出力するプログラム

## プロジェクトの構成
- プロジェクトのディレクトリ構造
<pre>
└─Backlogcsv2xlsx
    │  App.config
    │  BacklogEntity.cs
    │  Define.cs
    │  Form1.cs
    │  LineXlsx.cs
    │  Programs.cs
    └─Packages.config

</pre>

## 技術スタック
- C#
- ClosedXML
- CSVHelper

## 環境構築
1. Visual Studio 2022C#開発環境をインストール

## 運用方法
### 手順
1. ターゲットランタイムが.NET 4.7.2になっていることを確認しビルドを実行します。ビルドが正常に完了すると、実行可能ファイルが生成されます。
2. 必要に応じてApp.configファイルを設定してください。
#### iniファイル設定項目
- DayOfWeekEnd:初期状態の開始日に設定される曜日
- CrossMeetingSpan:クロスミーティングが開催される期間
- SortKeyDepartment:部門のソートに使用されるファイルのディレクトリのパス
- SortKeyDigital:デジ推向け項目のソートに使用されるファイルのディレクトリのパス
- isOpenDialog:起動直後にダイアログを表示するなら"1"、それ以外なら"0"
- isOpenXlsx:実行後に生成されたファイルを自動で開くなら"1"、それ以外なら"0"
- isClose:実行後プログラムを自動で終了するなら"1"、残すのであれば"0"
