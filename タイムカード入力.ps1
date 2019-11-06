# --------------------------------------------------
# �e�L�X�g�t�@�C���̓��e���^�C���J�[�h�ɓ��͂���
# --------------------------------------------------

# URL���w�肵��IE�N��
$url = "https://d4c-lt.com/contents/samplepage/timecard.html"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# ���͑Ώۗv�f�i�J�n�����A�I�������j�̕ϐ��ݒ�
# �敪�̗v�f�͔z��Ŏ擾���邽�ߑΏۊO�ɂ��Ă���
$startTime = $doc.getElementsByName("start")
$endTime = $doc.getElementsByName("end")

# �J�n���b�Z�[�W
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("OK�{�^�������œ��͂��܂�","�����J�n")

# �e�L�X�g�t�@�C���̓ǂݍ���
$filePath = "TimeCard.txt"
$inputFile = (Get-Content $filePath) -as [string[]]

# ���͏����̊J�n
$i = 1
foreach ($readLine in $inputFile) {
    $col = $readLine.split("`t")

    # �敪�̓���
    if ($col[1] -eq "-") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "0"
    }
    elseif ($col[1] -eq "�o��") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "1"
    }
    elseif ($col[1] -eq "�x��") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "2"
    }
    elseif ($col[1] -eq "�N�x") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "3"
    }
    elseif ($col[1] -eq "����") {
        $doc.getElementsByName("kubun_" + $i)[0].value = "4"
    }

    # �J�n�����̓���
    $startTime[$i -1].value = $col[2]

    # �I�������̓���
    $endTime[$i -1].value = $col[3]

    $i += 1
}

# �������b�Z�[�W
[System.Windows.Forms.MessageBox]::Show("�������������܂���","��������")