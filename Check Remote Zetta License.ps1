<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机提示错误的问题
==            修复xlog文件里读取不到授权的报错
==            支持检查多个Zetta Debug目录
== 2020/05/08 修复5.20.1.617 xlog文件中加入计算机名称的问题
== 2020/05/09 修复5月9号，会读取4月29号的xlog文件
=============================================================
#>

$RCSFileCreateDate = "{0:yyyyMMdd}" -f (Get-Date)
$RCSFileName = "RCS Check "+$RCSFileCreateDate+".txt"
$ComputerList = @()
$ProviderNameList = @()
$ZettaLicenseServerList = @()
$ZettaInstanceList = @()
$GSInstanceList = @()
$ZettaDebugPathList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
    ZettaLicenseServer = "^\s*ZettaLicenseServer\s*$";    
    ZettaInstance = "`\s*ZettaInstance\s*$";
    GSInstance = "`\s*GSInstance\s*$";
    ZettaDebugPath = "^\s*ZettaDebugPath\s*$";
}

$UserName = "RCS-USER"
$Password = ConvertTo-SecureString "RCS" -AsPlainText -Force;
$Cred = New-Object System.Management.Automation.PSCredential($UserName,$Password)

Get-Content $Path\config.ini | Foreach-Object {
  
    $var = $_.Split('=')
    $settingsKeys.Keys |% {
        if ($var[0] -match $settingsKeys.Item($_))
        {
            if ($_ -eq 'ComputerName')
            {
                $ComputerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ProviderName')
            {
                $ProviderNameList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaLicenseServer')
            {
                $ZettaLicenseServerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaInstance')
            {
                $ZettaInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'GSInstance')
            {
                $GSInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaDebugPath')
            {
                $ZettaDebugPathList += $var[1].Trim()
            }
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$flag = $false

$ZettaLicenseResults = ""

foreach($ZettaLicenseServer in $ZettaLicenseServerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ZettaLicenseServer'"
    $PingResult = Get-WmiObject -query $PingQuery

    $ZettaLicenseServer
    $LicenseDate = ""
    $List = @()

    if($PingResult.ProtocolAddress){

    foreach($ZettaDebugPath in $ZettaDebugPathList){
#    $DebugPath = '\\'+$ZettaLicenseServer+'\c$\ProgramData\RCS\Zetta\!Logging\Debug'
        $ZettaDebugPathURL = '\\'+$ZettaLicenseServer+$ZettaDebugPath
        $ZettaDebugPathURL
        if($PingResult.__SERVER -like $ZettaLicenseServer){}
        else{
            net use \\$ZettaLicenseServer\c$ /USER:RCS-USER RCS /PERSISTENT:YES
            net use \\$ZettaLicenseServer\d$ /USER:RCS-USER RCS /PERSISTENT:YES
            }
        $Day = "{0:dd}" -f (Get-Date)
        $ZettaDebugList = Get-ChildItem $ZettaDebugPathURL -Recurse -Include *"$Day"*Zetta.StartupManager.exe* -Filter *.xlog
        $ZettaLicenseResult = "机器名称: $ZettaLicenseServer`n"
        if($ZettaDebugList){
            $ZettaDebugFile = $ZettaDebugList | Sort-Object LastAccessTime -Descending | Select-Object -Index 0 | Select-Object -Property Name
            $ZettaDebugFileName = $ZettaDebugFile.name
            $GetZettaLicenseDate = Get-Content $ZettaDebugPathURL\"$ZettaDebugFileName" | Select-String -Pattern "License expiration date"
            $GetZettaLicenseCount = $GetZettaLicenseDate.Count
            $GetZettaLicenseCount
            if($GetZettaLicenseCount -gt 0){

            Get-Content $ZettaDebugPathURL\"$ZettaDebugFileName" | Select-String -Pattern "License expiration date" | ForEach-Object {
            $List += $_.ToString().Substring(8,36)
            }

            $LicenseDate = $list[$List.Count-1]
            $DateResult = $LicenseDate.ToString().Substring(25)
            $DateResult = Get-Date $DateResult -Format 'yyyy/MM/dd'
            $RemainDays = (New-TimeSpan $(Get-Date) $DateResult).Days

            if($RemainDays -ge 14){
                $ZettaLicenseResult = $ZettaLicenseResult + "授权到期日期: "+$DateResult+", 还有"+$RemainDays+"天授权到期!`n"
                break
                }
            elseif(($RemainDays -lt 14) -and ($RemainDays -gt 0)){
                $ZettaLicenseResult = $ZettaLicenseResult + "授权到期日期: "+$DateResult+", 还有"+$RemainDays+"天授权到期!`n"
                $flag = $true
                break
                }
            else{
                $ZettaLicenseResult = $ZettaLicenseResult + "授权已过期, 请立即联系RCS工程师重新授权!`n"
                $flag = $true
                break
                }
            }
            else{
                $ZettaLicenseResult = $ZettaLicenseResult + "成功读取到Zetta日志, 但是没有检测到Zetta授权, 请检查Zetta日志目录设置!`n"
                $flag = $true
                }
            }
    
        else{
            $ZettaLicenseResult = $ZettaLicenseResult + "没有检查到Zetta授权, 请打开Zetta!`n"
            $flag = $true   
            }
        }
    }
    else{
        $ZettaLicenseResult = "$ZettaLicenseServer 离线!`n"
        $flag = $true
        }

$ZettaLicenseResults = $ZettaLicenseResults + $ZettaLicenseResult + "`n"

}

Write-Host "$ZettaLicenseResults"

#RCS Monitor Used

$CheckData.OutString =  "------详情------`n$ZettaLicenseResults" 
$CheckData.OutState = $flag