<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
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
$WarningComputers = ""

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
        
        $EventResult = "机器名称: $ComputerName`n"
        if($events){        
            foreach($event in $events){
                $EventTime = $event.TimeCreated
                $EventName = $event.ProviderName
                $EventLevel = $event.LevelDisplayName
                $EventMessage = $event.Message
                $Flag = $true
                $EventResult = $EventResult + "$EventTime $EventName $EventLevel`n"
                }
            }
        else{
            $EventResult = $EventResult + "正常`n"
            }
        $EventResults = $EventResults + $EventResult + "`n"
        }
}

Write-Host "$EventResults"
$Flag

#RCS Monitor Used
$CheckData.OutString =  "------详情------`n$EventResults" 
$CheckData.OutState = $Flag