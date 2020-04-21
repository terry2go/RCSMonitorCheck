<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
=============================================================
#>

$ComputerList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    ComputerName = "^\s*ComputerName\s*$";
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
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$Flag = $false
$LastBootTimeResults = ""
$WarningComputers = ""

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
        
        $LastBootTimeResult = "机器名称: $ComputerName`n"
        $WarningComputer = ""
        foreach($LastBootTime in $LastBootTimes)
        {
            $LastBootDate = $LastBootTime.ConvertToDateTime($LastBootTime.lastbootuptime)
            $LastBootTimeResult = $LastBootTimeResult + "上次开机时间:" + $LastBootDate + "`n"
            $DateResult = $LastBootDate.Date
            $BootTimeSpan = New-TimeSpan $DateResult $(Get-Date)
            Write-Host $ComputerName $BootTimeSpan.Days "天没有重启了"
            if($BootTimeSpan.Days -ge 30){
                $Flag = $true
                $WarningComputer = $WarningComputer + "$ComputerName 超过30天没有重启了!`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`n"
            }
        }
    }
    else{
        $LastBootTimeResult = "$ComputerName 离线!`n"
    }
    $LastBootTimeResults = $LastBootTimeResults + $LastBootTimeResult + "`n"
}

Write-Host "$LastBootTimeResults"
Write-Host "$WarningComputers"

#RCS Monitor Used
$CheckData.OutString =  "------警告------`n$WarningComputers`n------详情------`n$LastBootTimeResults" 
$CheckData.OutState = $Flag