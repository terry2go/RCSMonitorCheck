<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
=============================================================
#>

$LicenseDate = ""
$List = @()
$path = '\\TERRY-T450\c$\ProgramData\RCS\Zetta\!Logging\Debug'
$Day = $(Get-Date).Day
$ZettaDebugList = Get-ChildItem $path -Recurse -Include *"$Day"_Zetta.StartupManager.exe* 
$ZettaDebugFile = $ZettaDebugList | Sort-Object LastAccessTime -Descending | Select-Object -Index 0 | Select-Object -Property Name
$ZettaDebugFileName = $ZettaDebugFile.name

Get-Content $path\"$ZettaDebugFileName" | Select-String -Pattern "License expiration date" | ForEach-Object {
    $List += $_.ToString().Substring(8,36)
}

$LicenseDate = $list[$List.Count-1]
$DateResult = $LicenseDate.ToString().Substring(25)
$DateResult = Get-Date $DateResult -Format 'yyyy/MM/dd'
$RemainDays = (New-TimeSpan $(Get-Date) $DateResult).Days

if($RemainDays -ge 0){
    $LicenseResult = "授权到期日期: "+$DateResult+", 还有"+$RemainDays+"天授权到期!"
}
else{
    $LicenseResult = "授权已过期, 请立即联系RCS工程师重新授权!"
}


Write-Host "$LicenseResult"

#RCS Monitor Used

$CheckData.OutString =  "$LicenseResult" 
$CheckData.OutState = ($RemainDays -le 14)