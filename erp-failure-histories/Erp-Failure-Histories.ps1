<#
◎実行例
Powershell -ExecutionPolicy Bypass -File Erp-Failure-Histories.ps1

◎フィールドコードとフィールド名の対応表
　フィールドコード　　フィールド名
　-----------------+-------------------
　ドロップダウン_1　　システム名
　ドロップダウン_4　　作業ステータス
　日時_1　　　　　　　障害発生日時
　日時_2　　　　　　　完了日時
　文字列__1行__4　　　サーバ名
　リッチエディター　　障害内容
　リッチエディター_0　作業内容
　ドロップダウン_4　　作業ステータス

#>

Add-Type -Path .\Net45\HtmlAgilityPack.dll
. .\ConvertFrom-HTML.ps1 # HTMLタグを除去するために使う

Set-Variable -Scope Script -name headers -value @{'X-Cybozu-API-Token' = $Env:Cybozu_API_Token} -Option Constant
Set-Variable -Scope Script -name lastMonth -value (Get-Date).AddMonths(-1).ToString("yyyy-MM-01") -Option Constant
Set-Variable -Scope Script -name thisMonth -value (Get-Date).ToString("yyyy-MM-01") -Option Constant

$Script:headers | FL

# クエリ生成
$Script:queries = New-Object System.Collections.Generic.List[string]
$Script:queries.Add((([string]::Join(' and ', @(
    'ドロップダウン_1 in ("Ｒ３")', 
    'ドロップダウン_4 in ("未着手", "調査中", "")',
    ('日時_1 < ' + """${Script:lastMonth}""")
))) + ' order by 日時_1 asc limit 500'))

$Script:queries.Add((([string]::Join(' and ', @(
    'ドロップダウン_1 in ("Ｒ３")', 
    'ドロップダウン_4 in ("未着手", "調査中", "完了")',
    ('日時_1 >= ' + """${Script:lastMonth}"""),
    ('日時_1 < ' + """${Script:thisMonth}""")
))) + ' order by レコード番号 desc limit 500'))

$Script:queries | FL

# CSVファイルをすべて削除する
Remove-Item * -Include *.csv

# クエリの実行結果を保存する一時ファイル
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
    レスポンスに含めるフィールドコードを設定する。
    書式は下記のとおり。
    @{
        "fields[0]" = "レコード番号"
        "fields[1]" = "文字列__1行__4"
        "fields[2]" = "日時_1"
        "fields[3]" = "ユーザー選択"
        "fields[4]" = "リッチエディター"
        "fields[5]" = "ラジオボタン_0"
        "fields[6]" = "リッチエディター_0"
        "fields[7]" = "ドロップダウン_4"
        "fields[8]" = "日時_2"
    }
    #>
    @(
        'レコード番号',
        '文字列__1行__4',
        '日時_1',
        'ユーザー選択',
        'リッチエディター',
        'ラジオボタン_0',
        'リッチエディター_0',
        'ドロップダウン_4',
        '日時_2'
    ) | ForEach-Object -Begin {$counter = 0} -Process { $Script:payload["field[${counter}]"] = $_ ;$counter++}
    Start-Sleep -Seconds 3
    # kintone APIにてデータ取得
    Invoke-WebRequest -Method Get -Uri $Env:BASE_URL -OutFile $file -Headers $Script:headers -Body $Script:payload 
}

# CSVファイルを併合する
Get-Content -Encodin UTF8 -Raw $Script:tempFiles | `
ConvertFrom-Json | `
Select-Object -ExpandProperty records | `
Select-Object @{Name='レコード番号';Expression={$_."レコード番号".value}}, `
              @{Name='サーバ名';Expression={$_."文字列__1行__4".value}}, `
              @{Name='障害発生日時';Expression={(Get-Date $_."日時_1".value).ToString("yyyy-MM-dd")}}, `
              @{Name='対応者';Expression={$_."ユーザー選択".value.name}}, `
              @{Name='障害内容';Expression={(ConvertFrom-Html $_."リッチエディター".value).innerText}}, `
              @{Name='重要度';Expression={$_."ラジオボタン_0".value}}, `
              @{Name='作業内容';Expression={(ConvertFrom-Html $_."リッチエディター_0".value).innerText}}, `
              @{Name='作業ステータス';Expression={$_."ドロップダウン_4".value}}, `
              @{Name='対処完了日時';Expression={if($_."日時_2".value){(Get-Date $_."日時_2".value).ToString("yyyy-MM-dd")}else{""}}} | `                  
ConvertTo-CSV -NoTypeInformation | `
Set-Content $Env:OUTPUT
