<#
=============================================================
== Terry Li
== 2020/04/20 初始版本
== 2020/04/21 增加GS授权检测
==            增加操作系统版本检测
== 2020/04/22 修复Win7下运行结果的显示

== 计算机检查清单：
== 1.检查操作系统版本
== 2.检查硬盘剩余空间
== 3.检查网卡状态
== 4.检查机器开机时间
== 5.检查系统日志

== RCS软件检查清单：
== 1.检查Zetta数据库
== 2.检查GS数据库
== 3.检查Zetta授权
== 4.检查GS授权

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

"
RCS Check Result
" | Out-File d:\$RCSFileName

"
计算机检查
" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 1.检查操作系统版本
=============================================================
" | Out-File -Append d:\$RCSFileName

$BuildVersionResults = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $OperatingSystems = Get-WmiObject -class Win32_OperatingSystem -computername $ComputerName
    }
    else{
        $OperatingSystems = Get-WmiObject -class Win32_OperatingSystem -Credential $Cred -computername $ComputerName
    }
        
        $BuildVersionResult = "机器名称: $ComputerName`r`n"
        foreach($OperatingSystem in $OperatingSystems)
        {
            $OperatingSystemCaption = $OperatingSystem.Caption
            $OperatingSystemBuildNumber = $OperatingSystem.BuildNumber
            $BuildVersionResult = $BuildVersionResult + "操作系统: " + $OperatingSystemCaption + ", 版本号: " + $OperatingSystemBuildNumber +"`r`n"

        }
    }
    else{
        $BuildVersionResult = "$ComputerName 离线!`r`n"
    }
    $BuildVersionResults = $BuildVersionResults + $BuildVersionResult + "`r`n"
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$BuildVersionResults" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 2.检查硬盘剩余空间
=============================================================
" | Out-File -Append d:\$RCSFileName
$DiskResults = ""
$WarningComputersDisk = ""
foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $Disks = Get-WmiObject -Class win32_logicaldisk -computername $ComputerName -filter "DriveType=3"
    }
    else{
        $Disks = Get-WmiObject -Class win32_logicaldisk -Credential $Cred -computername $ComputerName -filter "DriveType=3"
        }
        $DiskResult = "机器名称: $ComputerName`r`n"
        foreach($Disk in $Disks)
        {
            $DiskSize = " {0:0.0} GB" -f ($Disk.Size/1GB)
            $DiskFreeSize = " {0:0.0} GB" -f ($Disk.FreeSpace/1GB)
            $DiskDeviceID = $Disk.DeviceID
            $DiskVolumeName = $Disk.VolumeName
            $DiskPercentFree = [math]::Round((1-($Disk.FreeSpace/$Disk.Size))*100, 2)
            $DiskResult = $DiskResult + "$DiskDeviceID $DiskVolumeName`r`n剩余空间:$DiskFreeSize $DiskPercentFree%" + "`r`n"
            $WarningComputerDisk = "机器名称: $ComputerName`r`n"
            if($Disk.FreeSpace/1GB -lt 10)
            {
                $Flag = $true
                $WarningComputerDisk = $WarningComputerDisk + "$DiskDeviceID 盘剩余空间不足10GB, 请注意!`r`n"
                $WarningComputersDisk = $WarningComputersDisk + $WarningComputerDisk + "`r`n"
             }
        }
    }
    else{
        $DiskResult = "$ComputerName 离线!`r`n"
    }
    $DiskResults = $DiskResults + $DiskResult + "`r`n"
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$DiskResults" | Out-File -Append d:\$RCSFileName
"
== 错误信息：
" | Out-File -Append d:\$RCSFileName
"$WarningComputersDisk" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 3.检查网卡状态
=============================================================
" | Out-File -Append d:\$RCSFileName

$NetAdapterResults = ""
#$NetAdaptersSpeedList = New-Object -TypeName System.Collections.ArrayList
$WarningComputersNet = ""
$NetAdapters=""
foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $NetAdapters = Get-WmiObject -class Win32_NetworkAdapter -computername $ComputerName -Filter "PhysicalAdapter=True and NetEnabled=True"
        }
    else{
        $NetAdapters = Get-WmiObject -class Win32_NetworkAdapter -Credential $Cred -computername $ComputerName -Filter "PhysicalAdapter=True and NetEnabled=True"
        }

        $NetAdapterResult = "机器名称: $ComputerName`r`n"
        
        foreach($NetAdapter in $NetAdapters)
        {
            $NetAdaptersSpeed = "{0:0.0} Gbps" -f ($NetAdapter.Speed/1000000000)
            $NetAdapterResult = $NetAdapterResult + "网卡: " + $NetAdapter.NetConnectionID + ", 速度: " + $NetAdaptersSpeed + "`r`n"
            $WarningComputerNet = "机器名称: $ComputerName`r`n"
            if($NetAdapter.Speed/1000000000 -lt 1)
            {
                $Flag = $true
                $WarningComputerNet = $WarningComputerNet + $NetAdapter.NetConnectionID + " 网卡速度小于1Gbps!`r`n"
                $WarningComputersNet = $WarningComputersNet + $WarningComputerNet + "`r`n"
            }
        }
    }
    else{
        $NetAdapterResult = "$ComputerName 离线!`r`n"
    }
    $NetAdapterResults = $NetAdapterResults + $NetAdapterResult + "`r`n"
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$NetAdapterResults" | Out-File -Append d:\$RCSFileName
"
== 错误信息：
" | Out-File -Append d:\$RCSFileName
"$WarningComputersNet" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 4.检查机器开机时间
=============================================================
" | Out-File -Append d:\$RCSFileName

$LastBootTimeResults = ""
$WarningComputersLastBoot = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $LastBootTimes = Get-WmiObject -class Win32_OperatingSystem -computername $ComputerName
    }
    else{
        $LastBootTimes = Get-WmiObject -class Win32_OperatingSystem -Credential $Cred -computername $ComputerName
    }
        
        $LastBootTimeResult = "机器名称: $ComputerName`r`n"
        $WarningComputerLastBoot = ""
        foreach($LastBootTime in $LastBootTimes)
        {
            $LastBootDate = $LastBootTime.ConvertToDateTime($LastBootTime.lastbootuptime)
            $LastBootTimeResult = $LastBootTimeResult + "上次开机时间:" + $LastBootDate + "`r`n"
            $DateResult = $LastBootDate.Date
            $BootTimeSpan = New-TimeSpan $DateResult $(Get-Date)
            Write-Host $ComputerName $BootTimeSpan.Days "天没有重启了"
            if($BootTimeSpan.Days -ge 30){
                $Flag = $true
                $WarningComputerLastBoot = $WarningComputerLastBoot + "$ComputerName 超过30天没有重启了!`r`n"
                $WarningComputersLastBoot = $WarningComputersLastBoot + $WarningComputerLastBoot + "`r`n"
            }
        }
    }
    else{
        $LastBootTimeResult = "$ComputerName 离线!`r`n"
    }
    $LastBootTimeResults = $LastBootTimeResults + $LastBootTimeResult + "`r`n"
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$LastBootTimeResults" | Out-File -Append d:\$RCSFileName
"
== 错误信息：
" | Out-File -Append d:\$RCSFileName
"$WarningComputersLastBoot" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 5.检查系统日志
=============================================================
" | Out-File -Append d:\$RCSFileName

$StartTime = [datetime]::today
$EndTime = [datetime]::now
$Flag = $false
$EventResults = ""
$WarningComputersEventLog = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

        $EventFilter = @{Logname='System','Application'
                 Level=2,3
                 StartTime=$StartTime
                 EndTime=$EndTime
                 ProviderName=$ProviderNameList
                 }

    if($PingResult.__SERVER -like $ComputerName){
        $events = Get-WinEvent -computername $ComputerName -FilterHashtable $EventFilter -MaxEvents 5
    }
    else{
        $events = Get-WinEvent -Credential $Cred -computername $ComputerName -FilterHashtable $EventFilter -MaxEvents 5
    }
        
        $EventResult = "机器名称: $ComputerName`r`n"
        if($events){        
            foreach($event in $events){
                $EventTime = $event.TimeCreated
                $EventName = $event.ProviderName
                $EventLevel = $event.LevelDisplayName
                $EventMessage = $event.Message
                $Flag = $true
                $EventResult = $EventResult + "$EventTime $EventName $EventLevel`r`n"
                }
            }
        else{
            $EventResult = $EventResult + "正常`r`n"
            }
        $EventResults = $EventResults + $EventResult + "`r`n"
        }
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$EventResults" | Out-File -Append d:\$RCSFileName

"
RCS软件检查
" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 1.检查Zetta数据库
=============================================================
" | Out-File -Append d:\$RCSFileName

$QueryResultsZetta=""
$WarningComputersZetta=""

foreach($ZettaInstance in $ZettaInstanceList){
    $ServerName=$ZettaInstance
    $DatabaseName="ZettaDB"
    $UserName="sa"
    $Password="12h2oSt"
    $QueryResultZetta ="机器名称: $ServerName`r`n"
    $Query="
    select name, convert(float,size) * (8192.0/1024.0)/1024 from dbo.sysfiles
    SELECT database_name,MAX(backup_finish_date) AS backup_finish_date  FROM msdb.dbo.backupset where database_name = 'ZettaDB' GROUP BY database_name
    "
    $Conn=New-Object System.Data.SqlClient.SQLConnection
    $ConnectionString = "Data Source=$ServerName;Initial Catalog=$DatabaseName;user id=$UserName;pwd=$Password"
    $Conn.ConnectionString=$ConnectionString
    try{
        $Conn.Open()
        Write-Host "已连上数据库."
        
    }
    catch [exception]{
        Write-Warning "无法连接数据库."
        $Conn.Dispose()
        $flag=$true
        $WarningComputer= "无法连接 " + $ServerName + " 数据库." + "`r`n"
        $WarningComputersZetta = $WarningComputersZetta + $WarningComputerZetta + "`r`n"
        break
    }
    $SqlCommand=New-Object system.Data.SqlClient.SqlCommand($Query,$Conn)
    $DataSet=New-Object system.Data.DataSet
    $SqlDataAdapter=New-Object system.Data.SqlClient.SqlDataAdapter($SqlCommand)
    [void]$SqlDataAdapter.fill($DataSet)
#    $DataSet.Tables | fl
    $1 = $DataSet.Tables.database_name
    $2 = $DataSet.Tables.backup_finish_date
    $3 = $DataSet.Tables.name
    $4 = $DataSet.Tables.Column1
    $DBResults=""
    $WarningComputerZetta = "机器名称: $ServerName`r`n"
    for($i=0;$i -lt $3.count;$i++){
        if($3[$i]){
            $DBResult = $3[$i] + "的文件大小为: " + $4[$i] + "MB"
            $DBResults = $DBResults + $DBResult + "`r`n"
            if($4[$i] -gt 4096){
                $flag=$true
                $WarningComputerZetta = $WarningComputerZetta + $3[$i] + "大小超过4GB."+ "`r`n"
                $WarningComputersZetta = $WarningComputersZetta + $WarningComputerZetta + "`r`n"
            }
            }
        }
    $ServerName | out-file -Append d:\$RCSFileName
    $DataSet.Tables | fl | out-file -Append d:\$RCSFileName     
    $QueryResultZetta =  $QueryResultZetta + $DatabaseName + "的最新备份时间为: " + $2 + "`r`n" + $DBResults + "`r`n"
    $QueryResultsZetta = $QueryResultsZetta + $QueryResultZetta + "`r`n"
    $Conn.Close()
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
$QueryResultsZetta | Out-File -Append d:\$RCSFileName
"
== 错误信息：
" | Out-File -Append d:\$RCSFileName
$WarningComputersZetta | Out-File -Append d:\$RCSFileName

"
=============================================================
== 2.检查GS数据库
=============================================================
" | Out-File -Append d:\$RCSFileName

$QueryResultsGS=""
$WarningComputersGS=""

foreach($GSInstance in $GSInstanceList){
    $ServerName=$GSInstance
    $DatabaseName="gs"
    $UserName="sa"
    $Password="12h2oSt"
    $DBResults=""
    $WarningComputerGS = "机器名称: $ServerName`r`n"
    $QueryResultGS="机器名称: $ServerName`r`n"
    $Query="
    select name, convert(float,size) * (8192.0/1024.0)/1024 from dbo.sysfiles
    SELECT database_name,MAX(backup_finish_date) AS backup_finish_date  FROM msdb.dbo.backupset where database_name = 'gs' GROUP BY database_name
    "
    $Conn=New-Object System.Data.SqlClient.SQLConnection
    $ConnectionString = "Data Source=$ServerName;Initial Catalog=$DatabaseName;user id=$UserName;pwd=$Password"
    $Conn.ConnectionString=$ConnectionString
    try{
        $Conn.Open()
        Write-Host "已连上数据库."
        
    }
    catch [exception]{
        Write-Warning "无法连接数据库."
        $Conn.Dispose()
        $flag=$true
        $WarningComputerGS= "无法连接 " + $ServerName + " 数据库." + "`r`n"
        $WarningComputersGS = $WarningComputersGS + $WarningComputerGS + "`r`n"
        break
    }
    $SqlCommand=New-Object system.Data.SqlClient.SqlCommand($Query,$Conn)
    $DataSet=New-Object system.Data.DataSet
    $SqlDataAdapter=New-Object system.Data.SqlClient.SqlDataAdapter($SqlCommand)
    [void]$SqlDataAdapter.fill($DataSet)
    
    $1 = $DataSet.Tables.database_name
    $2 = $DataSet.Tables.backup_finish_date
    $3 = $DataSet.Tables.name
    $4 = $DataSet.Tables.Column1

    for($i=0;$i -lt $3.count;$i++){
        if($3[$i]){
            $DBResult = $3[$i] + "的文件大小为: " + $4[$i] + "MB"
            $DBResults = $DBResults + $DBResult + "`r`n"
            if($4[$i] -gt 4096){
                $flag=$true
                $WarningComputerGS = $WarningComputerGS + $3[$i] + "大小超过4GB."+ "`r`n"
                $WarningComputersGS = $WarningComputersGS + $WarningComputerGS + "`r`n"
            }
            }
        }
    $ServerName | out-file -Append d:\$RCSFileName
    $DataSet.Tables | fl | out-file -Append d:\$RCSFileName
    $QueryResultGS =  $QueryResultGS + $DatabaseName + "的最新备份时间为: " + $2 + "`r`n" + $DBResults + "`r`n"
    $QueryResultsGS = $QueryResultsGS + $QueryResultGS + "`r`n"
    $Conn.Close()
}

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
$QueryResultsGS | Out-File -Append d:\$RCSFileName
"
== 错误信息：
" | Out-File -Append d:\$RCSFileName
$WarningComputersGS | Out-File -Append d:\$RCSFileName

"
=============================================================
== 3.检查Zetta授权
=============================================================
" | Out-File -Append d:\$RCSFileName

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
        if($PingResult.__SERVER -like $ComputerName){}
        else{
            net use \\$ZettaLicenseServer\c$ /USER:RCS-USER RCS /PERSISTENT:YES
            net use \\$ZettaLicenseServer\d$ /USER:RCS-USER RCS /PERSISTENT:YES
            }
        $Day = $(Get-Date).Day
        $ZettaDebugList = Get-ChildItem $ZettaDebugPathURL -Recurse -Include *"$Day"_Zetta.StartupManager.exe* -Filter *.xlog
        $ZettaLicenseResult = "机器名称: $ZettaLicenseServer`r`n"
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
                $ZettaLicenseResult = $ZettaLicenseResult + "授权到期日期: "+$DateResult+", 还有"+$RemainDays+"天授权到期!`r`n"
                break
                }
            elseif(($RemainDays -lt 14) -and ($RemainDays -gt 0)){
                $ZettaLicenseResult = $ZettaLicenseResult + "授权到期日期: "+$DateResult+", 还有"+$RemainDays+"天授权到期!`r`n"
                $flag = $true
                break
                }
            else{
                $ZettaLicenseResult = $ZettaLicenseResult + "授权已过期, 请立即联系RCS工程师重新授权!`r`n"
                $flag = $true
                break
                }
            }
            else{
                $ZettaLicenseResult = $ZettaLicenseResult + "成功读取到Zetta日志, 但是没有检测到Zetta授权, 请检查Zetta日志目录设置!`r`n"
                $flag = $true
                }
            }
    
        else{
            $ZettaLicenseResult = $ZettaLicenseResult + "没有检查到Zetta授权, 请打开Zetta!`r`n"
            $flag = $true   
            }
        }
    }
    else{
        $ZettaLicenseResult = "$ZettaLicenseServer 离线!`r`n"
        $flag = $true
        }

$ZettaLicenseResults = $ZettaLicenseResults + $ZettaLicenseResult + "`r`n"

}


"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$ZettaLicenseResults" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 4.检查GS授权
=============================================================
" | Out-File -Append d:\$RCSFileName

$GSStationResults = ""
foreach($GSLicenseServer in $GSLicenseServerList){
    $PingQuery = "select * from win32_pingstatus where address = '$GSLicenseServer'"
    $PingResult = Get-WmiObject -query $PingQuery
    $GSStationResult = "机器名称: $GSLicenseServer`r`n"
    if($PingResult.ProtocolAddress){
        $gsUrl = "http://"+$GSLicenseServer+"/GSImportExportService/GSImportExportService.asmx"
        try{
        $GSWebProxy = New-WebServiceProxy -Uri $gsUrl -ErrorAction stop
        Write-Host "获取WebService成功!`r`n"
        }
        catch [Exception]{
        $GSStationResult = $GSStationResult + "获取WebService失败!`r`n"
        $GSStationResults = $GSStationResults + $GSStationResult + "`r`n"
        $flag = $true
        Write-Host "获取WebService失败!`r`n"
        continue
        }

        [Xml]$xmlVal = $GSWebProxy.GetStations()
        $Stations = $xmlVal.GSelector.Station
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

"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$GSStationResults" | Out-File -Append d:\$RCSFileName

