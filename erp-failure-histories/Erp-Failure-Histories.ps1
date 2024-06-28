<#
�����s��
Powershell -ExecutionPolicy Bypass -File Erp-Failure-Histories.ps1

���t�B�[���h�R�[�h�ƃt�B�[���h���̑Ή��\
�@�t�B�[���h�R�[�h�@�@�t�B�[���h��
�@-----------------+-------------------
�@�h���b�v�_�E��_1�@�@�V�X�e����
�@�h���b�v�_�E��_4�@�@��ƃX�e�[�^�X
�@����_1�@�@�@�@�@�@�@��Q��������
�@����_2�@�@�@�@�@�@�@��������
�@������__1�s__4�@�@�@�T�[�o��
�@���b�`�G�f�B�^�[�@�@��Q���e
�@���b�`�G�f�B�^�[_0�@��Ɠ��e
�@�h���b�v�_�E��_4�@�@��ƃX�e�[�^�X

#>

Add-Type -Path .\Net45\HtmlAgilityPack.dll
. .\ConvertFrom-HTML.ps1 # HTML�^�O���������邽�߂Ɏg��

Set-Variable -Scope Script -name headers -value @{'X-Cybozu-API-Token' = $Env:Cybozu_API_Token} -Option Constant
Set-Variable -Scope Script -name lastMonth -value (Get-Date).AddMonths(-1).ToString("yyyy-MM-01") -Option Constant
Set-Variable -Scope Script -name thisMonth -value (Get-Date).ToString("yyyy-MM-01") -Option Constant

$Script:headers | FL

# �N�G������
$Script:queries = New-Object System.Collections.Generic.List[string]
$Script:queries.Add((([string]::Join(' and ', @(
    '�h���b�v�_�E��_1 in ("�q�R")', 
    '�h���b�v�_�E��_4 in ("������", "������", "")',
    ('����_1 < ' + """${Script:lastMonth}""")
))) + ' order by ����_1 asc limit 500'))

$Script:queries.Add((([string]::Join(' and ', @(
    '�h���b�v�_�E��_1 in ("�q�R")', 
    '�h���b�v�_�E��_4 in ("������", "������", "����")',
    ('����_1 >= ' + """${Script:lastMonth}"""),
    ('����_1 < ' + """${Script:thisMonth}""")
))) + ' order by ���R�[�h�ԍ� desc limit 500'))

$Script:queries | FL

# CSV�t�@�C�������ׂč폜����
Remove-Item * -Include *.csv

# �N�G���̎��s���ʂ�ۑ�����ꎞ�t�@�C��
$Script:tempFiles = 1..3 | foreach {[System.IO.Path]::GetTempFileName()}

$i = 0
foreach ($q in $Script:queries) {
    $file = $Script:tempFiles[($i++)]
    Write-Host $file

    $Script:payload = @{
        "app" = $Env:APP_ID
        "query" = $q
        "totalCount" = "true"
    }
    <#
    ���X�|���X�Ɋ܂߂�t�B�[���h�R�[�h��ݒ肷��B
    �����͉��L�̂Ƃ���B
    @{
        "fields[0]" = "���R�[�h�ԍ�"
        "fields[1]" = "������__1�s__4"
        "fields[2]" = "����_1"
        "fields[3]" = "���[�U�[�I��"
        "fields[4]" = "���b�`�G�f�B�^�["
        "fields[5]" = "���W�I�{�^��_0"
        "fields[6]" = "���b�`�G�f�B�^�[_0"
        "fields[7]" = "�h���b�v�_�E��_4"
        "fields[8]" = "����_2"
    }
    #>
    @(
        '���R�[�h�ԍ�',
        '������__1�s__4',
        '����_1',
        '���[�U�[�I��',
        '���b�`�G�f�B�^�[',
        '���W�I�{�^��_0',
        '���b�`�G�f�B�^�[_0',
        '�h���b�v�_�E��_4',
        '����_2'
    ) | ForEach-Object -Begin {$counter = 0} -Process { $Script:payload["field[${counter}]"] = $_ ;$counter++}
    Start-Sleep -Seconds 3
    # kintone API�ɂăf�[�^�擾
    Invoke-WebRequest -Method Get -Uri $Env:BASE_URL -OutFile $file -Headers $Script:headers -Body $Script:payload 
}

# CSV�t�@�C���𕹍�����
Get-Content -Encodin UTF8 -Raw $Script:tempFiles | `
ConvertFrom-Json | `
Select-Object -ExpandProperty records | `
Select-Object @{Name='���R�[�h�ԍ�';Expression={$_."���R�[�h�ԍ�".value}}, `
              @{Name='�T�[�o��';Expression={$_."������__1�s__4".value}}, `
              @{Name='��Q��������';Expression={(Get-Date $_."����_1".value).ToString("yyyy-MM-dd")}}, `
              @{Name='�Ή���';Expression={$_."���[�U�[�I��".value.name}}, `
              @{Name='��Q���e';Expression={(ConvertFrom-Html $_."���b�`�G�f�B�^�[".value).innerText}}, `
              @{Name='�d�v�x';Expression={$_."���W�I�{�^��_0".value}}, `
              @{Name='��Ɠ��e';Expression={(ConvertFrom-Html $_."���b�`�G�f�B�^�[_0".value).innerText}}, `
              @{Name='��ƃX�e�[�^�X';Expression={$_."�h���b�v�_�E��_4".value}}, `
              @{Name='�Ώ���������';Expression={if($_."����_2".value){(Get-Date $_."����_2".value).ToString("yyyy-MM-dd")}else{""}}} | `                  
ConvertTo-CSV -NoTypeInformation | `
Set-Content $Env:OUTPUT
