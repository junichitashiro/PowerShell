# --------------------------------------------------
# テキストファイルの内容をタイムカードに入力する
# --------------------------------------------------

# URLを指定してIE起動
$url = "https://d4c-lt.com/contents/samplepage/timecard.html"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# 入力対象要素（開始時刻、終了時刻）の変数設定
# 区分の要素は配列で取得するため対象外にしている
$startTime = $doc.getElementsByName("start")
$endTime = $doc.getElementsByName("end")

# 開始メッセージ
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("OKボタン押下で入力します","処理開始")

# テキストファイルの読み込み
$filePath = "TimeCard.txt"
$inputFile = (Get-Content $filePath) -as [string[]]

# 入力処理の開始
$i = 1
foreach ($readLine in $inputFile) {
    $col = $readLine.split("`t")

    # 区分の入力
    if ($col[1] -eq "-") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "0"
    }
    elseif ($col[1] -eq "出勤") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "1"
    }
    elseif ($col[1] -eq "休日") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "2"
    }
    elseif ($col[1] -eq "年休") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "3"
    }
    elseif ($col[1] -eq "欠勤") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "4"
    }

    # 開始時刻の入力
    $startTime[$i -1].value = $col[2]

    # 終了時刻の入力
    $endTime[$i -1].value = $col[3]

    $i += 1
}

# 完了メッセージ
[System.Windows.Forms.MessageBox]::Show("処理が完了しました","処理完了")