<#
=============================================================
== Terry Li
== 2020/04/21 初始版本
== 2020/04/22 修复多数据库的授权显示问题
== 2020/04/30 增加软件版本检查
== 2020/05/09 加入“可以获取GS WebService，但是无法调用函数”的错误提示
=============================================================
#>

$RCSFileCreateDate = "{0:yyyyMMdd}" -f (Get-Date)
$RCSFileName = "RCS Check "+$RCSFileCreateDate+".txt"
$ComputerList = @()
$ProviderNameList = @()
$ZettaLicenseServerList = @()
$GSLicenseServerList = @()
$ZettaInstanceList = @()
$GSInstanceList = @()
$ZettaDebugPathList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
    ZettaLicenseServer = "^\s*ZettaLicenseServer\s*$";
    GSLicenseServer = "^\s*GSLicenseServer\s*$";     
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
            elseif ($_ -eq 'GSLicenseServer')
            {
                $GSLicenseServerList += $var[1].Trim()
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
$GSStationResults = ""
foreach($GSLicenseServer in $GSLicenseServerList){
    $PingQuery = "select * from win32_pingstatus where address = '$GSLicenseServer'"
    $PingResult = Get-WmiObject -query $PingQuery
    $GSStationResult = "机器名称: $GSLicenseServer`r`n"
    if($PingResult.ProtocolAddress){
        $gsUrl = "http://"+$GSLicenseServer+"/GSImportExportService/GSImportExportService.asmx"
        try{
        $GSWebProxy = New-WebServiceProxy -Uri $gsUrl -ErrorAction stop
        Write-Host "$GSLicenseServer 获取WebService成功!`r`n"
        }
        catch [Exception]{
        $GSStationResult = $GSStationResult + "获取WebService失败!`r`n"
        $GSStationResults = $GSStationResults + $GSStationResult + "`r`n"
        $flag = $true
        Write-Host "$GSLicenseServer 获取WebService失败!`r`n"
        continue
        }

        try{
            [Xml]$xmlVal = $GSWebProxy.GetStations()
            }
        catch [exception]{
            $GSStationResult = $GSStationResult + "获取电台信息失败!`r`n"
            $GSStationResults = $GSStationResults + $GSStationResult + "`r`n"
            $flag = $true
            Write-Host "$GSLicenseServer 获取电台信息失败!`r`n"
            continue
            }
        $Stations = $xmlVal.GSelector.Station
        $FirstStationID = $Stations[0].internalID
        [Xml]$xmlVal2 = $GSWebProxy.GetSystemInfo($FirstStationID)
        $GSVersion = $xmlVal2.SystemInfo.version
        $GSStationResult = $GSStationResult + "GS版本为: $GSVersion`r`n"
        foreach($Station in $Stations){
            $StationName = $Station.name
            $StationProducts = $Station.Products
            $StationExpiryDate = $Station.ExpiryDate
            try{
            $RemainDays = (New-TimeSpan $(Get-Date) $StationExpiryDate).Days
            }
            catch [exception]{
            $GSStationResult = $GSStationResult + "没有检查到GS授权!`r`n"
            $flag = $true
            break
            }
            if($RemainDays -gt 14){
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权到期日期: " + $StationExpiryDate + ", 还有" + $RemainDays + "天授权到期!`r`n"
                }
            elseif(($RemainDays -lt 14) -and ($RemainDays -gt 0)){
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权到期日期: " + $StationExpiryDate + ", 还有" + $RemainDays + "天授权到期!`r`n"
                $flag = $true
                }
            else{
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权已过期, 请立即联系RCS工程师重新授权!`r`n"
                $flag = $true
                }
            }
        }
    else{
        $GSStationResult = $GSStationResult + "机器离线!`r`n"
        $flag = $true
    }  
$GSStationResults = $GSStationResults + $GSStationResult + "`r`n"    
}

Write-Host $GSStationResults

#RCS Monitor Used

$CheckData.OutString =  "------详情------`r`n$GSStationResults" 
$CheckData.OutState = $flag