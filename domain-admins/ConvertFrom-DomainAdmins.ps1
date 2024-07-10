<#
監査資料のAD特権ユーザ一覧を生成するデータソースのデータクリーニング
#>
# powershell -ExecutionPolicy Bypass -NoProfile -File ConvertFrom-DomainAdmins.ps1

# 定数
Set-Variable -Scope Script -name Cache -value .\cache.txt -Option Constant
Set-Variable -Scope Script -name DataSourceOfLastMonth -value (Get-Content $script:Cache) -Option Constant

try {
    # 先月と今月の差異を特定するために使用する先月のアカウント一覧
    $script:lastMonthUser = Get-Content $script:DataSourceOfLastMonth | ConvertFrom-CSV | Select-Object @{Name='アカウント';Expression={$_.cn}}
    Write-Host $script:lastMonthUser
} catch {
    Write-Host $_
}

# 元ファイルをバックアップする。
$script:bkup = "$([System.IO.Path]::GetFileNameWithoutExtension($Env:DATASOURCE))_$((Get-Date).ToString("yyyyMMdd"))$([System.IO.Path]::GetExtension($Env:DATASOURCE))"
Copy-Item $Env:DATASOURCE -Destination $script:bkup
Set-Content -Path $script:Cache -Value $script:bkup

# CSVを目的の形に変換する
$script:csv = Get-Content $Env:DATASOURCE | ConvertFrom-CSV
$script:csv | `
	Select-Object @{Name='前月との差分.';Expression={if($script:lastMonthUser.'アカウント'.contains($_.cn)){""}else{"×"} }}, `
				  @{Name='St.';Expression={if([convert]::ToString($_."userAccountControl ", 16).EndsWith("2")) {'Lock'} else {''}}}, `
				  @{Name='稟議書No.';Expression={''}}, `
				  @{Name='アカウント';Expression={$_.cn}}, `
				  @{Name='氏名';Expression={$_.sn + '　' + $_.givenName}} | `
	ConvertTo-Csv -NoTypeInformation | `
	Select -Skip 1 | `
    % {$_.Replace('"','')} | `
	Out-File -Encoding default $Env:OUTPUT
