# VBA-Renew-Modules
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

# 使い方

## 設定
実行サンプル「実行サンプル モジュール自動更新.xlsm」の中の設定は以下の通り。


### 設定1(Excelの設定)

Excelの設定でExcel2019の場合「Excelのオプション」→「トラストセンター」→「マクロの設定」で

「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れておく
![Excelの設定](https://user-images.githubusercontent.com/73621859/126287884-57db4a75-3f34-4b35-b23d-f705067a1869.jpg)


### 設定2（使用モジュール）

-  ModRenewModules.bas


### 設定3（参照ライブラリ）

- Microsoft Visual Basic for Applications Extensibility 5.3  (VBAコードをVBAで参照するため)


![階層化フォーム 参照ライブラリ](https://user-images.githubusercontent.com/73621859/128787617-59d52e7e-0439-4f6c-9877-4bfe11e8d745.jpg)


実行環境など報告していただくと感謝感激雨霰。


## 使用例

### 起動時イベントでの実行（Workbook_Open）

### 更新するモジュールの設定
ワークブックの同じフォルダ上に更新対象のモジュールを置いておく

### 使用例システム