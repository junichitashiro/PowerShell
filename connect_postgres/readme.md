# 処理概要

PowerShellからPostgreSQLのDBに接続する

## 処理内容

1. PowerShellからPostgreSQLに接続する
2. スクリプトに記載したSQLを実行する
3. 実行結果を外部ファイルに出力する

### 対象テーブル名

* test_table（[サンプルテーブルの作成](https://github.com/junichitashiro/Knowledges/blob/master/DB/PostgreSQL/サンプルテーブルの作成.md) 参照）

### 実行SQL

* **$dbCmd.CommandText** に格納する

### 出力ファイル

* **C:\temp** フォルダにCSVファイルで出力する

### 補足

* odbcドライバがない場合は下記から入手する
* https://www.postgresql.org/ftp/odbc/versions/msi/