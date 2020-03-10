# --------------------------------------------------
# �^�C���J�[�h�̓��e���e�L�X�g�t�@�C���ɏo�͂���
# --------------------------------------------------
# ���b�Z�[�W�o�͗p
Add-Type -AssemblyName System.Windows.Forms

# �o�̓t�@�C���p�X
$filePath = "TimeCard.txt"

# URL���w�肵��IE�N��
$url = "https://d4c-lt.com/contents/samplepage/timecard.html"
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# �J�n���b�Z�[�W
[System.Windows.Forms.MessageBox]::Show("OK�{�^�������œ��e���o�͂��܂�","�����J�n")

# �o�͗p�t�@�C����V�K�쐬
$OutputText = $null
Write-Output $OutputText | Set-Content $filePath -Encoding Default

# ���t�J�E���g�p
$i = 1
while ($i -le 31) {

  # �v�f"kubun_XX"�̂����I������Ă���q�v�f��text���擾����
  $kubun = $doc.getElementsByName("kubun_" + $i)
  $childNo = $kubun[0].selectedIndex
  $outKubun = $kubun[0].children[$childNo].text

  # �J�n���Ԃ̓��͂��Ȃ�������f�t�H���g�l��ݒ肷��
  $startTime = $doc.getElementsByName("start")
  if ($null -eq $startTime[$i -1].value) {
    $outStartTime = "09:00"
  }
  else {
    $outStartTime = $startTime[$i -1].value
  }

  # �I�����Ԃ̓��͂��Ȃ�������f�t�H���g�l��ݒ肷��
  $endTime = $doc.getElementsByName("end")
  if ($null -eq $endTime[$i -1].value) {
    $outEndTime = "17:30"
  }
  else {
    $outEndTime = $endTime[$i -1].value
  }

  # �^�u��؂�œ��e���o�͂���
  $OutputText = [String]$i + "`t" + $outKubun + "`t" + $outStartTime + "`t" + $outEndTime
  Write-Output $OutputText | Add-Content $filePath -Encoding Default

  $i += 1
}

# �������b�Z�[�W
[System.Windows.Forms.MessageBox]::Show("�������������܂���","��������")