<#
=============================================================
== Terry Li
== 2020/04/16 初始版本
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
$NetAdapterResults = ""
#$NetAdaptersSpeedList = New-Object -TypeName System.Collections.ArrayList
$WarningComputers = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){
        $NetAdapterResult = "$ComputerName 在线!`n"
        $NetAdapterResults = $NetAdapterResults + $NetAdapterResult + "`n"
    }
    else{
        $flag = $true
        $WarningComputer = "$ComputerName 离线!`n"
        $WarningComputers = $WarningComputers + $WarningComputer + "`n"
    }
}

Write-Host "$NetAdapterResults"
Write-Host "$WarningComputers"


$CheckData.OutString =  "------警告------`n$WarningComputers`n------详情------`n$NetAdapterResults" 
$CheckData.OutState = $flag