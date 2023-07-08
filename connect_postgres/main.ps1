# --------------------------------------------------
# DB接続パラメータの設定
# --------------------------------------------------
$serverName = "localhost"
$portNo = "5432"
$dbName = "postgres"
$userName = "postgres"
$password = "postgres"

# --------------------------------------------------
# DB接続処理の設定
# --------------------------------------------------
$dbConString = "Driver={PostgreSQL UNICODE};Server=$serverName;Port=$portNo;Database=$dbName;Uid=$userName;Pwd=$password;"
$dbCon = New-Object System.Data.Odbc.OdbcConnection
$dbCon.ConnectionString = $dbConString;
$dbCon.Open()

# --------------------------------------------------
# SQLコマンドの作成
# --------------------------------------------------
$dbCmd = $dbCon.CreateCommand();
$dbCmd.CommandText = "select * from test_table"

# --------------------------------------------------
# SQL実行結果をデータセットに格納する
# --------------------------------------------------
$dataAdp = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter($dbCmd)
$dataSet = New-Object -TypeName System.Data.DataSet
# 実行結果を破棄する
$dataAdp.Fill($dataSet) > $null

# --------------------------------------------------
# データセットを出力する
# --------------------------------------------------
$dataSet.Tables[0] | export-csv C:\temp\export_csvFile.csv -notypeinformation -Encoding Default

# --------------------------------------------------
# DBコネクションを閉じる
# --------------------------------------------------
$dbCon.Close()
