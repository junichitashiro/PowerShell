# ----------------------------------------------------------------------
# ��c��.COM��Web�T�C�g�����c�����ƏZ���𒊏o���ăe�L�X�g�ŏo�͂���
# ----------------------------------------------------------------------

# --------------------------------------------------
# �X�N���v�g�̏�������
# --------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# �o�̓t�@�C���p�X��ݒ肷��
$fileName = "D:\roomlist.txt"

# �o�̓t�@�C��������������
Write-Output "" > $fileName

# URL��ݒ肷��
$url = "https://www.kaigishitu.com/"

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

# --------------------------------------------------
# �S����������
# --------------------------------------------------
$doc.getElementById("submit_btn").click()
while($ie.Busy) { Start-Sleep -seconds 1 }

# --------------------------------------------------
# �f�[�^�擾����
# --------------------------------------------------
# �ȉ��͌������ʃy�[�W�Ɂu����ɕ\���v�{�^��������O��Ő݌v���Ă���
# �����擾��A�u����ɕ\���v�{�^��������ꍇ�N���b�N���ď����擾��������
$btnHeight = $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].offsetHeight
$i = 0
while($btnHeight -ne 0) {

    # �u����ɕ\���v�{�^���̑��݃`�F�b�N
    $btnHeight = $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].offsetHeight

    # �������ʌ����̊i�[�@10�����̑z��
    $resultCnt = $doc.getElementsByClassName("c-topics__heading p-buildinglist__heading").length -1

    while($i -le $resultCnt) {

        # ��c�������i�[����
        $room = $doc.getElementsByClassName("c-topics__heading p-buildinglist__heading")[$i].outerText

        # �Z������]���ȕ������������Ċi�[����
        $text = $doc.getElementsByClassName("p-buildinglist__access")[$i].outerText
        $address = $text.replace("�n�}������","")

        # �o�͕����񂩂���s���������Ă���
        $tmp = $room + "`t" + $address.replace("`r`n","")
        Write-Output $tmp >> $fileName

        $i += 1

        # �u����ɕ\���v�{�^���i�z��j���������牟��
        $doc.getElementsByClassName("listMore_btn btn-blue buildingMore")[0].click()
        while($ie.Busy) { Start-Sleep -seconds 3 }

    }
}
[System.Windows.Forms.Messagebox]::Show("�������������܂���","����")
