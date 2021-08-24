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

![ライブラリ参照](https://user-images.githubusercontent.com/73621859/130536891-a4018575-902b-4d4b-8ac7-1a280affa583.jpg)

実行環境など報告していただくと感謝感激雨霰。


## 使用例

### 起動時イベントでの実行（Workbook_Open）
![起動時イベント設定](https://user-images.githubusercontent.com/73621859/130537080-ac101693-b4d0-47c6-a4bc-d8313d5a48d7.jpg)


### 開発ユーザーのみ実行の設定と、更新するモジュールの設定
![ユーザー設定と更新モジュール設定](https://user-images.githubusercontent.com/73621859/130536951-02bab051-af57-4cf6-a0b0-b1a09e9d2a7f.jpg)

### 更新するモジュールの設定
ワークブックの同じフォルダ上に更新対象のモジュールを置いておく
![更新モジュールの場所](https://user-images.githubusercontent.com/73621859/130537363-6c1271f4-2a81-46ca-bd8d-119420215ba4.jpg)

### 使用例システム

![Github自動更新](https://user-images.githubusercontent.com/73621859/130537158-fecfebb7-430b-4b89-8bbe-8ec4f0ac4a35.jpg)

