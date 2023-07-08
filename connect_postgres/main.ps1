# --------------------------------------------------
# DB�ڑ��p�����[�^�̐ݒ�
# --------------------------------------------------
$serverName = "localhost"
$portNo = "5432"
$dbName = "postgres"
$userName = "postgres"
$password = "postgres"

# --------------------------------------------------
# DB�ڑ������̐ݒ�
# --------------------------------------------------
$dbConString = "Driver={PostgreSQL UNICODE};Server=$serverName;Port=$portNo;Database=$dbName;Uid=$userName;Pwd=$password;"
$dbCon = New-Object System.Data.Odbc.OdbcConnection
$dbCon.ConnectionString = $dbConString;
$dbCon.Open()

# --------------------------------------------------
# SQL�R�}���h�̍쐬
# --------------------------------------------------
$dbCmd = $dbCon.CreateCommand();
$dbCmd.CommandText = "select * from test_table"

# --------------------------------------------------
# SQL���s���ʂ��f�[�^�Z�b�g�Ɋi�[����
# --------------------------------------------------
$dataAdp = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter($dbCmd)
$dataSet = New-Object -TypeName System.Data.DataSet
# ���s���ʂ�j������
$dataAdp.Fill($dataSet) > $null

# --------------------------------------------------
# �f�[�^�Z�b�g���o�͂���
# --------------------------------------------------
$dataSet.Tables[0] | export-csv C:\temp\export_csvFile.csv -notypeinformation -Encoding Default

# --------------------------------------------------
# DB�R�l�N�V���������
# --------------------------------------------------
$dbCon.Close()
