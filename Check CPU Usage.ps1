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

$CPUResults = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $CPUs = Get-WmiObject -Class win32_processor -computername $ComputerName
    }
    else{
        $CPUs = Get-WmiObject -Class win32_processor -Credential $Cred -computername $ComputerName
        }
        $CPUResult = "机器名称: $ComputerName`n"
        foreach($CPU in $CPUs){
        $FreeCPU = "{0:0}%" -f $CPU.LoadPercentage
        $CPUResult = $CPUResult + "CPU 使用率: $FreeCPU" + "`n"
        }
    }
    else{
        $CPUResult = "$ComputerName 离线!`n"
    }
    $CPUResults = $CPUResults + $CPUResult + "`n"
}

Write-Host "$CPUResults"

$CheckData.OutString =  "------详情------`n$CPUResults" 
$CheckData.OutState = ((Get-Date).Seconds -gt 30)