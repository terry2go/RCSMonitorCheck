<#
=============================================================
== Author: Terry Li
== Version: 0.2
== 2020/04/20 初始版本
== 2020/04/21 增加GS授权检测
==            增加操作系统版本检测
== 2020/04/22 修复Win7下运行结果的显示
== 2020/04/30 修复运行过程中的错误信息显示
==            增加软件版本检查
== 2020/05/08 修复5.20.1.617 xlog文件中加入计算机名称的问题
==            增加系统配置检查：机器序列号、操作系统版本、硬盘信息、内存信息、CPU信息、显卡信息
== 2020/05/09 修复“5月9号，检查Zetta授权时会读取4月29号的xlog文件”的问题
==            加入“可以获取GS WebService，但是无法调用函数”的错误提示
==            修复无法连接Zetta数据库的错误提醒
== 2020/06/18 加入忽略运行过程中的错误
== 2020/07/01 检查结果加入RCS Check版本显示
==            修复Team网卡显示的问题

== 计算机检查清单：
== 1.检查操作系统版本
== 2.检查硬盘剩余空间
== 3.检查网卡状态
== 4.检查机器开机时间
== 5.检查系统日志
== 6.检查系统配置

== RCS软件检查清单：
== 1.检查Zetta数据库, 软件版本
== 2.检查GS数据库
== 3.检查Zetta授权
== 4.检查GS授权, 软件版本

=============================================================
#>
$ErrorActionPreference= "silentlycontinue"

$RCSFileCreateDate = "{0:yyyyMMdd}" -f (Get-Date)
$RCSFileName = "RCS Check v0.2 "+$RCSFileCreateDate+".txt"
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

# 检查系统配置：机器序列号、操作系统版本、硬盘信息、内存信息、CPU信息、显卡信息
function Check-SystemInfo {
    [CmdletBinding(SupportsShouldProcess=$false)]
    Param(
        [Parameter(Position=0, Mandatory=$false)] [PSCredential]$Credential,
        [Parameter(Position=1, Mandatory=$true)] [string]$Computername
        )
    Process {
        $UserName = "RCS-USER"
        $Password = ConvertTo-SecureString "RCS" -AsPlainText -Force;
        $Cred = New-Object System.Management.Automation.PSCredential($UserName,$Password)
        $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
        $PingResult = Get-WmiObject -query $PingQuery
        if($PingResult.ProtocolAddress){
            if($PingResult.__SERVER -like $ComputerName){
                $SystemInfo_Bios = Get-WmiObject -Class win32_bios -ComputerName $ComputerName
                $SystemInfo_OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
                $SystemInfo_Memory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $ComputerName
                $SystemInfo_Processor = Get-WmiObject -Class Win32_Processor -ComputerName $ComputerName
                $SystemInfo_Disk = Get-WmiObject -class Win32_DiskDrive -ComputerName $ComputerName
                $SystemInfo_Video = Get-WmiObject -Class CIM_VideoController -ComputerName $ComputerName
                }
            else{
                $SystemInfo_Bios = Get-WmiObject -Class win32_bios -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_OS = Get-WmiObject -Class Win32_OperatingSystem -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Memory = Get-WmiObject -Class Win32_PhysicalMemory -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Processor = Get-WmiObject -Class Win32_Processor -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Disk = Get-WmiObject -class Win32_DiskDrive -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Video = Get-WmiObject -Class CIM_VideoController -Credential $Cred -ComputerName $ComputerName
                }
            Write-Host "== "$ComputerName" 配置`r`n"
            $SystemInfo_Bios_Num = $SystemInfo_Bios.serialnumber
            $SystemInfo_Bios_Num_Result = "机器序列号: " + $SystemInfo_Bios_Num + "`r`n"
            Write-Host $SystemInfo_Bios_Num_Result
            $SystemInfo_OS_Caption = $SystemInfo_OS.caption
            $SystemInfo_OS_Build = $SystemInfo_OS.BuildNumber
            $SystemInfo_OS_Result = "操作系统: " + $SystemInfo_OS_Caption + " Build " + $SystemInfo_OS_Build + "`r`n"
            Write-Host $SystemInfo_OS_Result
            Write-Host "`r`n== 内存信息`r`n"
            $SystemInfo_Memory | Select-Object Manufacturer,SerialNumber,Capacity | ForEach-Object{
                $SystemInfo_Memory_Result = "内存品牌: " + $_.Manufacturer + ", 序列号: " + $_.SerialNumber + ", 容量: " + $_.Capacity/1GB + "G"
                $SystemInfo_Memory_Result2 = $SystemInfo_Memory_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Memory_Result2
                $SystemInfo_Memory_Results = $SystemInfo_Memory_Results + $SystemInfo_Memory_Result2 + "`r`n"
                }
            Write-Host "`r`n== CPU信息`r`n"
            $SystemInfo_Processor | Select-Object Manufacturer,Name,NumberOfCores,NumberOfLogicalProcessors | ForEach-Object{
                $SystemInfo_Processor_Result = "CPU: " + $_.Name + ", " + $_.NumberOfCores + "核, " + $_.NumberOfLogicalProcessors + "线程"
                $SystemInfo_Processor_Result2 = $SystemInfo_Processor_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Processor_Result2
                $SystemInfo_Processor_Results = $SystemInfo_Processor_Results + $SystemInfo_Processor_Result2 + "`r`n"
                }
            Write-Host "`r`n== 硬盘信息`r`n"
            $SystemInfo_Disk | Select-Object Model,SerialNumber,Size | Foreach-Object{
                $SystemInfo_Disk_Result = "硬盘: " + $_.model + ", 序列号: " + $_.SerialNumber + ", 容量: " + [math]::Round($_.size/1GB) + "GB"
                $SystemInfo_Disk_Result2 = $SystemInfo_Disk_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Disk_Result2
                $SystemInfo_Disk_Results = $SystemInfo_Disk_Results + $SystemInfo_Disk_Result2 + "`r`n"
                }
            Write-Host "`r`n== 显卡信息`r`n"
            $SystemInfo_Video | select Caption | ForEach-Object{
                $SystemInfo_Video_Result = "显卡: " + $_.Caption
                Write-Host $SystemInfo_Video_Result
                $SystemInfo_Video_Results = $SystemInfo_Video_Results + $SystemInfo_Video_Result + "`r`n"
                }
            $SystemInfo_Result = "== " + $ComputerName + " 配置`r`n`r`n" + $SystemInfo_Bios_Num_Result + "`r`n" + $SystemInfo_OS_Result + "`r`n" + $SystemInfo_Memory_Results + "`r`n"  + $SystemInfo_Processor_Results+ "`r`n"  + $SystemInfo_Disk_Results+ "`r`n"   + $SystemInfo_Video_Results
            $SystemInfo_Result
        }
    }
}


"
RCS Check Result
" | Out-File d:\$RCSFileName

Write-Host "
RCS Check Result
"

"
第一项：计算机检查
=============================================================
== 1.检查操作系统版本
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
第一项：计算机检查
=============================================================
== 1.检查操作系统版本
=============================================================
"

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

Write-Host "
== 检查结果：
$BuildVersionResults
"

"
== 检查结果：
$BuildVersionResults" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 2.检查硬盘剩余空间
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 2.检查硬盘剩余空间
=============================================================
"

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

Write-Host "
== 检查结果：
$DiskResults
== 错误信息：
$WarningComputersDisk
"

"
== 检查结果：
$DiskResults
== 错误信息：
$WarningComputersDisk" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 3.检查网卡状态
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 3.检查网卡状态
=============================================================
" 

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

        $NetAdapterResult = "机器名称: $ComputerName`n"
        $NetFlag = $false
        foreach($NetAdapter in $NetAdapters){
            if($NetAdapter.PNPDeviceID -notlike "*PCI*" -and $NetAdapter.ServiceName -notlike "*VBox*" -and $NetAdapter.ServiceName -notlike "*VM*"){
                $NetFlag = $True
                Write-Host "$ComputerName 的网卡有组Team."
            }
        }
        if($NetFlag){
        $NetAdapters = $NetAdapters | Where-Object {$_.PNPDeviceID -NotLike "*PCI*"}
        foreach($NetAdapter in $NetAdapters)
        {
            $NetAdaptersSpeed = "{0:0.0} Gbps" -f ($NetAdapter.Speed/1000000000)
            $NetAdapterResult = $NetAdapterResult + "网卡: " + $NetAdapter.NetConnectionID + ", 速度: " + $NetAdaptersSpeed + "`n"
            $WarningComputer = "机器名称: $ComputerName`n"
            if($NetAdapter.Speed/1000000000 -lt 1)
            {
                $Flag = $true
                $WarningComputer = $WarningComputer + $NetAdapter.NetConnectionID + " 网卡速度小于1Gbps!`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`n"
            }
        }
        }else{
        $NetAdapters = $NetAdapters | Where-Object {$_.PNPDeviceID -Like "*PCI*"}
        foreach($NetAdapter in $NetAdapters)
        {
            $NetAdaptersSpeed = "{0:0.0} Gbps" -f ($NetAdapter.Speed/1000000000)
            $NetAdapterResult = $NetAdapterResult + "网卡: " + $NetAdapter.NetConnectionID + ", 速度: " + $NetAdaptersSpeed + "`n"
            $WarningComputer = "机器名称: $ComputerName`n"
            if($NetAdapter.Speed/1000000000 -lt 1)
            {
                $Flag = $true
                $WarningComputer = $WarningComputer + $NetAdapter.NetConnectionID + " 网卡速度小于1Gbps!`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`n"
            }
        }
        }
    }
    else{
        $NetAdapterResult = "$ComputerName 离线!`n"
    }
    $NetAdapterResults = $NetAdapterResults + $NetAdapterResult + "`n"
}

Write-Host "
== 检查结果：
$NetAdapterResults
== 错误信息：
$WarningComputersNet
"

"
== 检查结果：
$NetAdapterResults
== 错误信息：
$WarningComputersNet" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 4.检查机器开机时间
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 4.检查机器开机时间
=============================================================
"

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
            $DateResult = $LastBootDate.Date
            $BootTimeSpan = New-TimeSpan $DateResult $(Get-Date)
            $LastBootTimeResult = $LastBootTimeResult + "上次开机时间:" + $LastBootDate + "`r`n" + $BootTimeSpan.Days + "天没有重启了.`r`n"
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

Write-Host "
== 检查结果：
$LastBootTimeResults
== 错误信息：
$WarningComputersLastBoot
"

"
== 检查结果：
$LastBootTimeResults
== 错误信息：
$WarningComputersLastBoot" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 5.检查系统日志
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 5.检查系统日志
=============================================================
"

$StartTime = [datetime]::today
$EndTime = [datetime]::now
$Flag = $false
$EventResults = ""
$WarningComputersEventLogs = ""

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
        
        try{
        $events = Get-WinEvent -computername $ComputerName -FilterHashtable $EventFilter -MaxEvents 5 -ErrorAction SilentlyContinue
        }
        catch [exception]
        {
        Write-Host "$ComputerName 发现异常, 无法检查系统日志."
        $WarningComputersEventLog = ""
        $WarningComputersEventLog = $WarningComputersEventLog + "$ComputerName 发现异常, 无法检查系统日志.`r`n"
        $Flag = $true
        continue
        }
    }
    else{
        try{
        $events = Get-WinEvent -Credential $Cred -computername $ComputerName -FilterHashtable $EventFilter -MaxEvents 5 -ErrorAction SilentlyContinue
        }
        catch [exception]
        {
        Write-Host "$ComputerName 发现异常, 无法检查系统日志."
        $WarningComputersEventLog = ""
        $WarningComputersEventLog = $WarningComputersEventLog + "$ComputerName 发现异常, 无法检查系统日志.`r`n"
        $Flag = $true
        continue
        }
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

Write-Host "
== 检查结果：
$EventResults
== 错误信息:
$WarningComputersEventLog
"

"
== 检查结果：
$EventResults
== 错误信息:
$WarningComputersEventLog" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 6.检查系统配置
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 6.检查系统配置
=============================================================
"

foreach ($ComputerName in $ComputerList){

    Check-SystemInfo -Computername $ComputerName | Out-File -Append d:\$RCSFileName

}

"
第二项：RCS软件检查
" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 1.检查Zetta数据库, 软件版本
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
第二项：RCS软件检查
=============================================================
== 1.检查Zetta数据库, 软件版本
=============================================================
"

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
    use ZettaDB
    select top 1 AppVersion from cpu
    "
    $Conn=New-Object System.Data.SqlClient.SQLConnection
    $ConnectionString = "Data Source=$ServerName;Initial Catalog=$DatabaseName;user id=$UserName;pwd=$Password"
    $Conn.ConnectionString=$ConnectionString
    try{
        $Conn.Open()
        Write-Host "$ServerName 已连上数据库."
        
    }
    catch [exception]{
        Write-Warning "$ServerName 无法连接数据库."
        $Conn.Dispose()
        $flag=$true
        $WarningComputer= "无法连接 " + $ServerName + " 数据库." + "`r`n"
        $WarningComputersZetta = $WarningComputersZetta + $WarningComputer + "`r`n"
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
    $5 = $DataSet.Tables.appversion
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
#    $ServerName
#    $DataSet.Tables | fl
    $QueryResultZetta =  $QueryResultZetta + "Zetta版本为: " + $5 + "`r`n" + $DatabaseName + "的最新备份时间为: " + $2 + "`r`n" + $DBResults + "`r`n"
    $QueryResultsZetta = $QueryResultsZetta + $QueryResultZetta + "`r`n"
    $Conn.Close()
}

Write-Host "
== 检查结果：
$QueryResultsZetta
== 错误信息：
$WarningComputersZetta
"

"
== 检查结果：
$QueryResultsZetta
== 错误信息：
$WarningComputersZetta" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 2.检查GS数据库
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 2.检查GS数据库
=============================================================
"

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
        Write-Host "$ServerName 已连上数据库."
        
    }
    catch [exception]{
        Write-Warning "$ServerName 无法连接数据库."
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

Write-Host "
== 检查结果：
$QueryResultsGS
== 错误信息：
$WarningComputersGS
"

"
== 检查结果：
$QueryResultsGS
== 错误信息：
$WarningComputersGS" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 3.检查Zetta授权
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 3.检查Zetta授权
=============================================================
"

$ZettaLicenseResults = ""

foreach($ZettaLicenseServer in $ZettaLicenseServerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ZettaLicenseServer'"
    $PingResult = Get-WmiObject -query $PingQuery
    $LicenseDate = ""
    $List = @()
    if($PingResult.ProtocolAddress){
        foreach($ZettaDebugPath in $ZettaDebugPathList){
        $ZettaDebugPathURL = '\\'+$ZettaLicenseServer+$ZettaDebugPath
        if($PingResult.__SERVER -like $ZettaLicenseServer){
            Write-Host "授权机器 $ZettaLicenseServer 为本机, 无需账户密码.`r`n"
        }
        else{
            net use \\$ZettaLicenseServer\c$ /USER:RCS-USER RCS /PERSISTENT:YES
            net use \\$ZettaLicenseServer\d$ /USER:RCS-USER RCS /PERSISTENT:YES
            Write-Host "授权机器 $ZettaLicenseServer 成功连接文件目录.`r`n"
            }
        $Day = "{0:dd}" -f (Get-Date)
        $ZettaDebugList = ""
        if(Test-Path $ZettaDebugPathURL){
            $ZettaDebugList = Get-ChildItem $ZettaDebugPathURL -Recurse -Include *"$Day"*Zetta.StartupManager.exe* -Filter *.xlog
        }
        
        $ZettaLicenseResult = "机器名称: $ZettaLicenseServer`r`n"
        if($ZettaDebugList){
            $ZettaDebugFile = $ZettaDebugList | Sort-Object LastAccessTime -Descending | Select-Object -Index 0 | Select-Object -Property Name
            $ZettaDebugFileName = $ZettaDebugFile.name
            $GetZettaLicenseDate = Get-Content $ZettaDebugPathURL\"$ZettaDebugFileName" | Select-String -Pattern "License expiration date"
            $GetZettaLicenseCount = $GetZettaLicenseDate.Count
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

Write-Host "
== 检查结果：
$ZettaLicenseResults
"


"
== 检查结果：
" | Out-File -Append d:\$RCSFileName
"$ZettaLicenseResults" | Out-File -Append d:\$RCSFileName

"
=============================================================
== 4.检查GS授权, 软件版本
=============================================================
" | Out-File -Append d:\$RCSFileName

Write-Host "
=============================================================
== 4.检查GS授权, 软件版本
=============================================================
"

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

Write-Host "
== 检查结果：
$GSStationResults
"

"
== 检查结果：
$GSStationResults" | Out-File -Append d:\$RCSFileName