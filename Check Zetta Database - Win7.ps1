<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 加入无法连接数据库时的错误警告
== 2020/04/30 增加软件版本检查
== 2020/05/09 修复无法连接Zetta数据库的错误提醒
== 2020/04/14 支持Win7
=============================================================
#>
$TempFile = "{0:MMddHHmmss}" -f (Get-Date)
$ComputerList = @()
$ProviderNameList = @()
$ZettaInstanceList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
    ZettaLicenseServer = "^\s*ZettaLicenseServer\s*$";
    ZettaInstance = "`\s*ZettaInstance\s*$";
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
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$QueryResultsZetta=""
$WarningComputersZetta=""
$flag=$flase

foreach($ZettaInstance in $ZettaInstanceList){
    $ServerName=$ZettaInstance
    $DatabaseName="ZettaDB"
    $UserName="sa"
    $Password="12h2oSt"
    $QueryResultZetta ="机器名称: $ServerName`r`n"
    $Query="
    use ZettaDB
    select name, convert(float,size) * (8192.0/1024.0)/1024 size from dbo.sysfiles
    SELECT database_name,MAX(backup_finish_date) AS backup_finish_date  FROM msdb.dbo.backupset where database_name = 'ZettaDB' GROUP BY database_name
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
    $QueryData = @{}
    $namelist = @()
    $sizelist = @()
    $DataSet.Tables | fl > $Path\$TempFile.txt
    $test = Get-Content -Encoding UTF8 $Path\$TempFile.txt | Where-Object {$_ -like '* : *'} | ForEach-Object {

        $infos = $_ -split ' : '
        if($infos[0] -eq "name"){
        $namelist += $infos[1].Trim() 
        }
        elseif($infos[0] -eq "size"){
        $sizelist += $infos[1].Trim()
        }
        else{
        $key = $infos[0].Trim()
        $value = $infos[1].Trim()
        }
        $QueryData.$key = $value  
    }
#    $DataSet.Tables | fl
    $1 = $QueryData.database_name
    $2 = $QueryData.backup_finish_date
    $3 = $namelist
    $4 = $sizelist
    $5 = $QueryData.appversion
    $DBResults=""
    $WarningComputersZetta = ""
    for($i=0;$i -lt $3.count;$i++){
        if($3[$i]){
            $WarningComputerZetta = ""
            $DBResult = $3[$i] + "的文件大小为: " + $4[$i] + "MB"
            $DBResults = $DBResults + $DBResult + "`r`n"
            if([Int]$4[$i] -gt 4096){
                $WarningComputerZetta = "机器名称: $ServerName`r`n"
                $flag=$true
                $WarningComputerZetta = $WarningComputerZetta + $3[$i] + "大小超过4GB." 
            }
            $WarningComputersZetta = $WarningComputersZetta + $WarningComputerZetta + "`r`n"
         }
            
    }

#    $ServerName | out-file -Append d:\$RCSFileName
#    $DataSet.Tables | fl | out-file -Append d:\$RCSFileName     
#    $ServerName
#    $DataSet.Tables | fl
    $QueryResultZetta =  $QueryResultZetta + "Zetta版本为: " + $5 + "`r`n" + $DatabaseName + "的最新备份时间为: " + $2 + "`r`n" + $DBResults + "`r`n"
    $QueryResultsZetta = $QueryResultsZetta + $QueryResultZetta + "`r`n"
    $Conn.Close()
    Remove-Item $Path\$TempFile.txt
}


Write-Host $QueryResultsZetta
Write-Host $WarningComputersZetta
#RCS Monitor Used

$CheckData.OutString =  "------警告------`n$WarningComputersZetta`n------详情------`n$QueryResultsZetta" 
$CheckData.OutState = $flag