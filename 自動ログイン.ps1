# --------------------------------------------------
# 事前に設定した内容で自動ログインする
# --------------------------------------------------

# 設定情報
$url = "https://d4c-lt.com/contents/samplepage/login1.html"
$id = "ID12345"
$pw = "PW12345"

# IEの起動から画面表示まで
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while ($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# ログイン情報の設定
$doc.getElementById("user_name").value = $id
$doc.getElementById("password").value = $pw
$doc.getElementById("login_btn").click()

# 処理確認用
# Add-Type -AssemblyName System.Windows.Forms
# [System.Windows.Forms.MessageBox]::Show("ログインを実行","処理確認用")