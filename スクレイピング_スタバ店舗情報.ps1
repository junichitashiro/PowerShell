# ----------------------------------------------------------------------
# �X�^�[�o�b�N�X��Web�T�C�g����X�ܖ��ƏZ���𒊏o���ăe�L�X�g�ŏo�͂���
# ----------------------------------------------------------------------

# --------------------------------------------------
# �X�N���v�g�̏�������
# --------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# �o�̓t�@�C���p�X��ݒ肷��
$fileName = "D:\shoplist.txt"

# �o�̓t�@�C��������������
Write-Output "" > $fileName

# URL��ݒ肷��
$url = "https://store.starbucks.co.jp/"

# --------------------------------------------------
# IE�̏�������
# --------------------------------------------------
# �����ݒ�URL�ֈړ�����
$ie = New-Object -ComObject InternetExplorer.Application

# ��ʂ̓������m�F�������ꍇ�͉��L���R�����g�C������
# $ie.Visible = $true
$ie.Navigate($url)

# �y�[�W���؂�ւ��܂ő҂�
while($ie.Busy) { Start-Sleep -seconds 1 }

# �h�L�������g�I�u�W�F�N�g���擾����
$doc = $ie.document

# �����N�{�^���v�f�ԍ��̊J�n�l�ƏI���l��ݒ肷��
# ��ӂɓ���ł����񂪂Ȃ����ߌŒ�l��ݒ肵�Ă���
$startLink = 131
$endLink   = 177

while($startLink -le $endLink) {

    # --------------------------------------------------
    # �s���{���̌�������
    # --------------------------------------------------
    # �s���{���̃����N�{�^�����N���b�N���Č��ʉ�ʂֈړ�����
    $doc.getElementsByTagName("a")[$startLink].click()
    while($ie.Busy) { Start-Sleep -seconds 3 }

    # --------------------------------------------------
    # �X�܏��̎擾����
    # --------------------------------------------------
    # �������ʉ�ʂŁu�����ƌ���v�{�^�����\������Ă����牟���邾������
    while($doc.getElementById("moreList").offsetLeft -ne 0) {
        $doc.getElementById("moreList").click()
        Start-Sleep -seconds 1
    }

    # �X�܏��v�f�ԍ��̊J�n�l�ƏI���l��ݒ肷��
    # ���������ӂɓ���ł����񂪂Ȃ����ߌŒ�l
    $tmp = ""
    $i = 41
    $maxLength = $doc.getElementsByTagName("P").length -14

    # --------------------------------------------------
    # �擾���̏o�͏���
    # --------------------------------------------------
    while ($i -le $maxLength) {
        $tmp = $doc.getElementsByTagName("P")[$i].outerText + "`t" + $doc.getElementsByTagName("P")[$i + 1].outerText
        # Write-Output $tmp
        Write-Output $tmp >> $fileName

        $i += 5
    }

    # �����ݒ�URL�֖߂�
    $ie.Navigate($url)
    while($ie.Busy) { Start-Sleep -seconds 3 }

    $startLink += 1
}
[System.Windows.Forms.Messagebox]::Show("�������������܂���","����")
