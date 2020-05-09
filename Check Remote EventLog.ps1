<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
== 2020/04/30 修复运行过程中的错误信息显示
=============================================================
#>

$ComputerList = @()
$ProviderNameList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
    ProviderName = "^\s*ProviderName\s*$";
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
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

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

#RCS Monitor Used
$CheckData.OutString =  "------警告------`n$WarningComputersEventLog`n------详情------`n$EventResults" 
$CheckData.OutState = $Flag