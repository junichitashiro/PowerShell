# --------------------------------------------------
# ���O�ɐݒ肵�����e�Ŏ������O�C������
# --------------------------------------------------

# �ݒ���
$url = "https://d4c-lt.com/contents/samplepage/login1.html"
$id = "ID12345"
$pw = "PW12345"

# IE�̋N�������ʕ\���܂�
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($url)
while ($ie.Busy) { Start-Sleep -milliseconds 100 }
$doc = $ie.document

# ���O�C�����̐ݒ�
$doc.getElementById("user_name").value = $id
$doc.getElementById("password").value = $pw
$doc.getElementById("login_btn").click()

# �����m�F�p
# Add-Type -AssemblyName System.Windows.Forms
# [System.Windows.Forms.MessageBox]::Show("���O�C�������s","�����m�F�p")