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

$MemoryResults = ""

Foreach($ComputerName in $ComputerList) {
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress) {

    if($PingResult.__SERVER -like $ComputerName){
        $Memorys = Get-WmiObject -Class win32_operatingsystem -ComputerName $ComputerName
    }
    else{
        $Memorys = Get-WmiObject -Class win32_operatingsystem -Credential $Cred -ComputerName $ComputerName
        }
        $MemoryResult = "机器名称: $ComputerName`n"
        foreach($Memory in $Memorys) {
        $FreeMemory = [math]::Round($Memory.FreePhysicalMemory / 1MB,2)
        $UsedMemory = [math]::Round(($Memory.TotalVisibleMemorySize-$Memory.FreePhysicalMemory) / 1MB,2)
        $TotalMemory = [math]::Round($Memory.TotalVisibleMemorySize / 1MB,2)
        $MemoryLoadPercentage = [math]::Round(($UsedMemory/$TotalMemory)*100, 0)
        $MemoryResult =  $MemoryResult + "内存使用: $UsedMemory / $TotalMemory GB ($MemoryLoadPercentage %)" + "`n"
        }
    }
    else{
        $MemoryResult = "$ComputerName 离线!`n"
    }
    $MemoryResults = $MemoryResults + $MemoryResult + "`n"
}

Write-Host "$MemoryResults"
$CheckData.OutString =  "------详情------`n$MemoryResults" 
$CheckData.OutState = ((Get-Date).Seconds -gt 30)