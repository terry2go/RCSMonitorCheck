<#
=============================================================
== Terry Li
== 2020/04/21 初始版本
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

#$StationList = @()
$flag = $false
foreach($GSLicenseServer in $GSLicenseServerList){
    $PingQuery = "select * from win32_pingstatus where address = '$GSLicenseServer'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){
        $gsUrl = "http://"+$GSLicenseServer+"/GSImportExportService/GSImportExportService.asmx"
        $GSWebProxy = New-WebServiceProxy -Uri $gsUrl
        [Xml]$xmlVal = $GSWebProxy.GetStations()
        $Stations = $xmlVal.GSelector.Station
        $GSStationResult = ""
        foreach($Station in $Stations){
            $StationName = $Station.name
            $StationProducts = $Station.Products
            $StationExpiryDate = $Station.ExpiryDate

            $RemainDays = (New-TimeSpan $(Get-Date) $StationExpiryDate).Days
            if($RemainDays -gt 14){
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权到期日期: " + $StationExpiryDate + ", 还有" + $RemainDays + "天授权到期!`n`n"
                }
            elseif(($RemainDays -lt 14) -and ($RemainDays -gt 0)){
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权到期日期: " + $StationExpiryDate + ", 还有" + $RemainDays + "天授权到期!`n`n"
                $flag = $true
            }
            else{
                $GSStationResult = $GSStationResult + "电台: " + $StationName + ", 授权已过期, 请立即联系RCS工程师重新授权!`n`n"
                $flag = $true 
            }
        }

        }
    }

Write-Host $GSStationResult

#RCS Monitor Used

$CheckData.OutString =  "------详情------`n$GSStationResult" 
$CheckData.OutState = $flag