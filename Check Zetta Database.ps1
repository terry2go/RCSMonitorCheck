<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 加入无法连接数据库时的错误警告
== 2020/04/30 增加软件版本检查
== 2020/05/09 修复无法连接Zetta数据库的错误提醒
== 2020/05/26 加入警告弹窗
== 2020/07/14 加入忽略运行过程中的错误
== 2020/07/27 将远程计算机的用户名以及密码单独写到ini配置文件，便于集中修改
==            加入自定义检测阈值
== 2020/07/29 修复一些异常
=============================================================
#>
$ErrorActionPreference = "silentlycontinue"
$ComputerList = @()
$ProviderNameList = @()
$ZettaInstanceList = @()
$Path = 'D:\Bat\RCSMonitor'
$MessageBox = [System.Windows.Forms.MessageBox]

$settingsKeys = @{
    UserName                 = "^\s*UserName\s*$";
    Password                 = "^\s*Password\s*$";
    FreeSpaceThreshold       = "^\s*FreeSpaceThreshold\s*$";
    DatabaseThreshold        = "^\s*DatabaseThreshold\s*$";
    BootTimeThreshold        = "^\s*BootTimeThreshold\s*$";
    NetAdapterSpeedThreshold = "^\s*NetAdapterSpeedThreshold\s*$";
    LicenseThreshold         = "^\s*LicenseThreshold\s*$";
    ComputerName             = "^\s*ComputerName\s*$";
    ProviderName             = "^\s*ProviderName\s*$";
    ZettaLicenseServer       = "^\s*ZettaLicenseServer\s*$";
    GSLicenseServer          = "^\s*GSLicenseServer\s*$";     
    ZettaInstance            = "`\s*ZettaInstance\s*$";
    GSInstance               = "`\s*GSInstance\s*$";
    ZettaDebugPath           = "^\s*ZettaDebugPath\s*$";
}

Get-Content $Path\config.ini | Foreach-Object {
    $var = $_.Split('=')
    $settingsKeys.Keys | % {
        if ($var[0] -match $settingsKeys.Item($_)) {
            if ($_ -eq 'ComputerName') {
                $ComputerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ProviderName') {
                $ProviderNameList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaLicenseServer') {
                $ZettaLicenseServerList += $var[1].Trim()
            }
            elseif ($_ -eq 'GSLicenseServer') {
                $GSLicenseServerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaInstance') {
                $ZettaInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'GSInstance') {
                $GSInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaDebugPath') {
                $ZettaDebugPathList += $var[1].Trim()
            }
            elseif ($_ -eq 'UserName') {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'Password') {
                $Password = $var[1].Trim()
            }
            elseif ($_ -eq 'UserName') {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'FreeSpaceThreshold') {
                $FreeSpaceThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'DatabaseThreshold') {
                $DatabaseThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'BootTimeThreshold') {
                $BootTimeThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'NetAdapterSpeedThreshold') {
                $NetAdapterSpeedThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'LicenseThreshold') {
                $LicenseThreshold = $var[1].Trim()
            }
            else {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$PasswordNew = ConvertTo-SecureString $Password -AsPlainText -Force;
$Cred = New-Object System.Management.Automation.PSCredential($UserName, $PasswordNew)

$QueryResultsZetta = ""
$WarningComputersZetta = ""
$flag = $flase

foreach ($ZettaInstance in $ZettaInstanceList) {
    $ServerName = $ZettaInstance
    $DatabaseName = "ZettaDB"
    $SQLUserName = "sa"
    $SQLPassword = "12h2oSt"
    $QueryResultZetta = "机器名称: $ServerName`r`n"
    $Query = "
    select name, convert(float,size) * (8192.0/1024.0)/1024 from dbo.sysfiles
    SELECT database_name,MAX(backup_finish_date) AS backup_finish_date  FROM msdb.dbo.backupset where database_name = '$DatabaseName' GROUP BY database_name
    use $DatabaseName
    select top 1 AppVersion from cpu
    "
    $Conn = New-Object System.Data.SqlClient.SQLConnection
    $ConnectionString = "Data Source=$ServerName;Initial Catalog=$DatabaseName;user id=$SQLUserName;pwd=$SQLPassword"
    $Conn.ConnectionString = $ConnectionString
    try {
        $Conn.Open()
        Write-Host "$ServerName 已连上数据库."
        
    }
    catch [exception] {
        Write-Warning "$ServerName 无法连接数据库."
        $Conn.Dispose()
        $flag = $true
        $WarningComputer = "无法连接 " + $ServerName + " 数据库." + "`r`n"
        $WarningComputersZetta = $WarningComputersZetta + $WarningComputer + "`r`n"
        break
    }
    $SqlCommand = New-Object system.Data.SqlClient.SqlCommand($Query, $Conn)
    $DataSet = New-Object system.Data.DataSet
    $SqlDataAdapter = New-Object system.Data.SqlClient.SqlDataAdapter($SqlCommand)
    [void]$SqlDataAdapter.fill($DataSet)
    #    $DataSet.Tables | fl
    $1 = $DataSet.Tables.database_name
    $2 = $DataSet.Tables.backup_finish_date
    $3 = $DataSet.Tables.name
    $4 = $DataSet.Tables.Column1
    $5 = $DataSet.Tables.appversion
    $DBResults = ""
    $WarningComputerZetta = ""
    for ($i = 0; $i -lt $3.count; $i++) {
        if ($3[$i]) {
            $DBResult = $3[$i] + "的文件大小为: " + $4[$i] + "MB"
            $DBResults = $DBResults + $DBResult + "`r`n"
            if ($4[$i] / 1024 -gt $DatabaseThreshold) {
                $flag = $true
                $WarningComputerZetta = "机器名称: $ServerName`r`n" + $3[$i] + "大小超过" + $DatabaseThreshold + "GB.`r`n"
            }
            else {
                $WarningComputerZetta = ""
            }
            $WarningComputersZetta = $WarningComputersZetta + $WarningComputerZetta + "`r`n"
        }
    }
    $ServerName
    $DataSet.Tables | fl
    $QueryResultZetta = $QueryResultZetta + "Zetta版本为: " + $5 + "`r`n" + $DatabaseName + "的最新备份时间为: " + $2 + "`r`n" + $DBResults + "`r`n"
    $QueryResultsZetta = $QueryResultsZetta + $QueryResultZetta + "`r`n"
    $Conn.Close()
}

Write-Host $QueryResultsZetta
Write-Host $flag
Write-Host $WarningComputersZetta
#RCS Monitor Used

if ($flag) {
    $MessageBox::Show("$WarningComputersZetta", "数据库警告")
}

$CheckData.OutString = "------警告------`n$WarningComputersZetta`n------详情------`n$QueryResultsZetta" 
$CheckData.OutState = $flag