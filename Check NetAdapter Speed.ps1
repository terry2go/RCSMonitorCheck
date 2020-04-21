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
$NetAdapterResults = ""
#$NetAdaptersSpeedList = New-Object -TypeName System.Collections.ArrayList
$WarningComputers = ""
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
    else{
        $NetAdapterResult = "$ComputerName 离线!`n"
    }
    $NetAdapterResults = $NetAdapterResults + $NetAdapterResult + "`n"
}

Write-Host "$NetAdapterResults"
Write-Host "$WarningComputers"


$CheckData.OutString =  "------警告------`n$WarningComputers`n------详情------`n$NetAdapterResults" 
$CheckData.OutState = $flag