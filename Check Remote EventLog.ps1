<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
== 2020/04/30 修复运行过程中的错误信息显示
== 2020/07/14 加入忽略运行过程中的错误
== 2020/07/27 将远程计算机的用户名以及密码单独写到ini配置文件，便于集中修改
=============================================================
#>
$ErrorActionPreference= "silentlycontinue"
$ComputerList = @()
$ProviderNameList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    UserName = "^\s*UserName\s*$";
    Password = "^\s*Password\s*$";
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
}

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
            elseif ($_ -eq 'UserName')
            {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'Password')
            {
                $Password = $var[1].Trim()
            }
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$PasswordNew = ConvertTo-SecureString $Password -AsPlainText -Force;
$Cred = New-Object System.Management.Automation.PSCredential($UserName,$PasswordNew)

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
        $WarningComputersEventLog = "$ComputerName 发现异常, 无法检查系统日志.`r`n"
        $WarningComputersEventLogs = $WarningComputersEventLog + $WarningComputersEventLogs
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
        $WarningComputersEventLog = "$ComputerName 发现异常, 无法检查系统日志.`r`n"
        $WarningComputersEventLogs = $WarningComputersEventLog + $WarningComputersEventLogs
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
                $WarningComputersEventLog = ""
                $WarningComputersEventLogs = $WarningComputersEventLog + "$ComputerName 发现错误信息, 请检查系统日志.`r`n"
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
$WarningComputersEventLogs
"

#RCS Monitor Used
$CheckData.OutString =  "------警告------`n$WarningComputersEventLogs`n------详情------`n$EventResults" 
$CheckData.OutState = $Flag