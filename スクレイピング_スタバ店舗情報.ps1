# ----------------------------------------------------------------------
# スターバックスのWebサイトから店舗名と住所を抽出してテキストで出力する
# ----------------------------------------------------------------------

# --------------------------------------------------
# スクリプトの初期処理
# --------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# 出力ファイルパスを設定する
$fileName = "D:\shoplist.txt"

# 出力ファイルを初期化する
Write-Output "" > $fileName

# URLを設定する
$url = "https://store.starbucks.co.jp/"

# --------------------------------------------------
# IEの初期処理
# --------------------------------------------------
# 初期設定URLへ移動する
$ie = New-Object -ComObject InternetExplorer.Application

# 画面の動きを確認したい場合は下記をコメントインする
# $ie.Visible = $true
$ie.Navigate($url)

# ページが切り替わるまで待つ
while($ie.Busy) { Start-Sleep -seconds 1 }

# ドキュメントオブジェクトを取得する
$doc = $ie.document

# リンクボタン要素番号の開始値と終了値を設定する
# 一意に特定できる情報がないため固定値を設定している
$startLink = 131
$endLink   = 177

while($startLink -le $endLink) {

    # --------------------------------------------------
    # 都道府県の検索処理
    # --------------------------------------------------
    # 都道府県のリンクボタンをクリックして結果画面へ移動する
    $doc.getElementsByTagName("a")[$startLink].click()
    while($ie.Busy) { Start-Sleep -seconds 3 }

    # --------------------------------------------------
    # 店舗情報の取得処理
    # --------------------------------------------------
    # 検索結果画面で「もっと見る」ボタンが表示されていたら押せるだけ押す
    while($doc.getElementById("moreList").offsetLeft -ne 0) {
        $doc.getElementById("moreList").click()
        Start-Sleep -seconds 1
    }

    # 店舗情報要素番号の開始値と終了値を設定する
    # こちらも一意に特定できる情報がないため固定値
    $tmp = ""
    $i = 41
    $maxLength = $doc.getElementsByTagName("P").length -14

    # --------------------------------------------------
    # 取得情報の出力処理
    # --------------------------------------------------
    while ($i -le $maxLength) {
        $tmp = $doc.getElementsByTagName("P")[$i].outerText + "`t" + $doc.getElementsByTagName("P")[$i + 1].outerText
        # Write-Output $tmp
        Write-Output $tmp >> $fileName

        $i += 5
    }

    # 初期設定URLへ戻る
    $ie.Navigate($url)
    while($ie.Busy) { Start-Sleep -seconds 3 }

    $startLink += 1
}
[System.Windows.Forms.Messagebox]::Show("処理が完了しました","完了")
