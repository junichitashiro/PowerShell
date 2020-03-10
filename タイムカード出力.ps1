# --------------------------------------------------
# タイムカードの内容をテキストファイルに出力する
# --------------------------------------------------
# メッセージ出力用
Add-Type -AssemblyName System.Windows.Forms

# 出力ファイルパス
$filePath = "TimeCard.txt"

# URLを指定してIE起動
$url = "https://d4c-lt.com/contents/samplepage/timecard.html"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# 開始メッセージ
[System.Windows.Forms.MessageBox]::Show("OKボタン押下で内容を出力します","処理開始")

# 出力用ファイルを新規作成
$OutputText = $null
Write-Output $OutputText | Set-Content $filePath -Encoding Default

# 日付カウント用
$i = 1
while ($i -le 31) {

  # 要素"kubun_XX"のうち選択されている子要素のtextを取得する
  $kubun = $doc.getElementsByName("kubun_" + $i)
  $childNo = $kubun[0].selectedIndex
  $outKubun = $kubun[0].children[$childNo].text

  # 開始時間の入力がなかったらデフォルト値を設定する
  $startTime = $doc.getElementsByName("start")
  if ($null -eq $startTime[$i -1].value) {
    $outStartTime = "09:00"
  }
  else {
    $outStartTime = $startTime[$i -1].value
  }

  # 終了時間の入力がなかったらデフォルト値を設定する
  $endTime = $doc.getElementsByName("end")
  if ($null -eq $endTime[$i -1].value) {
    $outEndTime = "17:30"
  }
  else {
    $outEndTime = $endTime[$i -1].value
  }

  # タブ区切りで内容を出力する
  $OutputText = [String]$i + "`t" + $outKubun + "`t" + $outStartTime + "`t" + $outEndTime
  Write-Output $OutputText | Add-Content $filePath -Encoding Default

  $i += 1
}

# 完了メッセージ
[System.Windows.Forms.MessageBox]::Show("処理が完了しました","処理完了")