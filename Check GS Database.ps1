<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 加入无法连接数据库时的错误警告
=============================================================
#>

$ComputerList = @()
$ProviderNameList = @()
$GSInstanceList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
    ZettaLicenseServer = "^\s*ZettaLicenseServer\s*$";
    GSInstance = "`\s*GSInstance\s*$";
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
            elseif ($_ -eq 'GSInstance')
            {
                $GSInstanceList += $var[1].Trim()
            }
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$QueryResults=""
$flag=$false
$WarningComputers=""

foreach($GSInstance in $GSInstanceList){
    $ServerName=$GSInstance
    $DatabaseName="gs"
    $UserName="sa"
    $Password="12h2oSt"
    $DBResults=""
    $WarningComputer = "机器名称: $ServerName`n"
    $QueryResult="机器名称: $ServerName`n"
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
        $WarningComputer= "无法连接 " + $ServerName + " 数据库." + "`n"
        $WarningComputers = $WarningComputers + $WarningComputer + "`n"
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

    for($i=0;$i -lt $3.count;$i++){
        if($3[$i]){
            $DBResult = $3[$i] + "的文件大小为: " + $4[$i] + "MB"
            $DBResults = $DBResults + $DBResult + "`n"
            if($4[$i] -gt 4096){
                $flag=$true
                $WarningComputer= $WarningComputer + $3[$i] + "大小超过4GB."+ "`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`n"
            }
            }
        }
         
    $QueryResult =  $QueryResult + $DatabaseName + "的最新备份时间为: " + $2 + "`n" + $DBResults + "`n"
    $QueryResults = $QueryResults + $QueryResult + "`n"
    $Conn.Close()
}


Write-Host $QueryResults
Write-Host $flag
Write-Host $WarningComputers
#RCS Monitor Used

$CheckData.OutString =  "------警告------`n$WarningComputers`n------详情------`n$QueryResults" 
$CheckData.OutState = $flag