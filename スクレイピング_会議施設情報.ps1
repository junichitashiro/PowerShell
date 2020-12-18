# ----------------------------------------------------------------------
# 会議室.COMのWebサイトから会議室名と住所を抽出してテキストで出力する
# ----------------------------------------------------------------------

# --------------------------------------------------
# スクリプトの初期処理
# --------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# 出力ファイルパスを設定する
$fileName = "D:\roomlist.txt"

# 出力ファイルを初期化する
Write-Output "" > $fileName

# URLを設定する
$url = "https://www.kaigishitu.com/"

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

# --------------------------------------------------
# 全件検索処理
# --------------------------------------------------
$doc.getElementById("submit_btn").click()
while($ie.Busy) { Start-Sleep -seconds 1 }

# --------------------------------------------------
# データ取得処理
# --------------------------------------------------
# 以下は検索結果ページに「さらに表示」ボタンがある前提で設計している
# 情報を取得後、「さらに表示」ボタンがある場合クリックして情報を取得し続ける
$btnHeight = $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].offsetHeight
$i = 0
while($btnHeight -ne 0) {

    # 「さらに表示」ボタンの存在チェック
    $btnHeight = $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].offsetHeight

    # 検索結果件数の格納　10件ずつの想定
    $resultCnt = $doc.getElementsByClassName("c-topics__heading p-buildinglist__heading").length -1

    while($i -le $resultCnt) {

        # 会議室名を格納する
        $room = $doc.getElementsByClassName("c-topics__heading p-buildinglist__heading")[$i].outerText

        # 住所から余分な文字を除去して格納する
        $text = $doc.getElementsByClassName("p-buildinglist__access")[$i].outerText
        $address = $text.replace("地図を見る","")

        # 出力文字列から改行を除去しておく
        $tmp = $room + "`t" + $address.replace("`r`n","")
        Write-Output $tmp >> $fileName

        $i += 1

        # 「さらに表示」ボタン（配列）があったら押す
        $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].click()
        while($ie.Busy) { Start-Sleep -seconds 3 }

    }
}
[System.Windows.Forms.Messagebox]::Show("処理が完了しました","完了")
