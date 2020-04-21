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
$DiskResults = ""
$WarningComputers = ""
foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $Disks = Get-WmiObject -Class win32_logicaldisk -computername $ComputerName -filter "DriveType=3"
    }
    else{
        $Disks = Get-WmiObject -Class win32_logicaldisk -Credential $Cred -computername $ComputerName -filter "DriveType=3"
        }
        $DiskResult = "机器名称: $ComputerName`n"
        foreach($Disk in $Disks)
        {
            $DiskSize = " {0:0.0} GB" -f ($Disk.Size/1GB)
            $DiskFreeSize = " {0:0.0} GB" -f ($Disk.FreeSpace/1GB)
            $DiskDeviceID = $Disk.DeviceID
            $DiskVolumeName = $Disk.VolumeName
            $DiskPercentFree = [math]::Round((1-($Disk.FreeSpace/$Disk.Size))*100, 2)
            $DiskResult = $DiskResult + "$DiskDeviceID $DiskVolumeName`n剩余空间:$DiskFreeSize $DiskPercentFree%" + "`n"
            $WarningComputer = "机器名称: $ComputerName`n"
            if($Disk.FreeSpace/1GB -lt 10)
            {
                $Flag = $true
                $WarningComputer = $WarningComputer + "$DiskDeviceID 盘剩余空间不足10GB, 请注意!`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`n"
             }
        }
    }
    else{
        $DiskResult = "$ComputerName 离线!`n"
    }
    $DiskResults = $DiskResults + $DiskResult + "`n"
}

Write-Host "$DiskResults"
Write-Host "$WarningComputers"

$CheckData.OutString =  "------警告------`n$WarningComputers`n------详情------`n$DiskResults" 
$CheckData.OutState = $Flag
