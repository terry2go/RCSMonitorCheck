<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
== 2020/05/28 加入警告弹窗
== 2020/06/10 修复Team网卡显示的问题
== 2020/07/14 加入忽略运行过程中的错误
== 2020/07/27 将远程计算机的用户名以及密码单独写到ini配置文件，便于集中修改
==            加入自定义检测阈值
=============================================================
#>
$ErrorActionPreference= "silentlycontinue"
$ComputerList = @()
$Path = 'D:\Bat\RCSMonitor'
$MessageBox = [System.Windows.Forms.MessageBox]

$settingsKeys = @{
    UserName                 = "^\s*UserName\s*$";
    Password                 = "^\s*Password\s*$";
    FreeSpaceThreshold       = "^\s*FreeSpaceThreshold\s*$";
    DatabaseThreshold        = "^\s*DatabaseThreshold\s*$";
    BootTimeThreshold        = "^\s*BootTimeThreshold\s*$";
    NetAdapterSpeedThreshold = "^\s*NetAdapterSpeedThreshold\s*$";
    LicenseThreshold         = "^\s*LicenseThreshold\s*$";
    ComputerName             = "^\s*ComputerName\s*$";
    ProviderName             = "^\s*ProviderName\s*$";
    ZettaLicenseServer       = "^\s*ZettaLicenseServer\s*$";
    GSLicenseServer          = "^\s*GSLicenseServer\s*$";     
    ZettaInstance            = "`\s*ZettaInstance\s*$";
    GSInstance               = "`\s*GSInstance\s*$";
    ZettaDebugPath           = "^\s*ZettaDebugPath\s*$";
}

Get-Content $Path\config.ini | Foreach-Object {
    $var = $_.Split('=')
    $settingsKeys.Keys | % {
        if ($var[0] -match $settingsKeys.Item($_)) {
            if ($_ -eq 'ComputerName') {
                $ComputerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ProviderName') {
                $ProviderNameList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaLicenseServer') {
                $ZettaLicenseServerList += $var[1].Trim()
            }
            elseif ($_ -eq 'GSLicenseServer') {
                $GSLicenseServerList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaInstance') {
                $ZettaInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'GSInstance') {
                $GSInstanceList += $var[1].Trim()
            }
            elseif ($_ -eq 'ZettaDebugPath') {
                $ZettaDebugPathList += $var[1].Trim()
            }
            elseif ($_ -eq 'UserName') {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'Password') {
                $Password = $var[1].Trim()
            }
            elseif ($_ -eq 'UserName') {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'FreeSpaceThreshold') {
                $FreeSpaceThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'DatabaseThreshold') {
                $DatabaseThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'BootTimeThreshold') {
                $BootTimeThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'NetAdapterSpeedThreshold') {
                $NetAdapterSpeedThreshold = $var[1].Trim()
            }
            elseif ($_ -eq 'LicenseThreshold') {
                $LicenseThreshold = $var[1].Trim()
            }
            else {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$PasswordNew = ConvertTo-SecureString $Password -AsPlainText -Force;
$Cred = New-Object System.Management.Automation.PSCredential($UserName, $PasswordNew)

$Flag = $false
$NetAdapterResults = ""
#$NetAdaptersSpeedList = New-Object -TypeName System.Collections.ArrayList
$WarningComputersNet = ""
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
        $NetFlag = $false
        foreach($NetAdapter in $NetAdapters){
            if($NetAdapter.PNPDeviceID -notlike "*PCI*" -and $NetAdapter.ServiceName -notlike "*VBox*" -and $NetAdapter.ServiceName -notlike "*VM*"){
                $NetFlag = $True
                Write-Host "$ComputerName 的网卡有组Team."
            }
        }
        if($NetFlag){
        $NetAdapters = $NetAdapters | Where-Object {$_.PNPDeviceID -NotLike "*PCI*"}
        foreach($NetAdapter in $NetAdapters)
        {
            $NetAdaptersSpeed = "{0:0.0} Gbps" -f ($NetAdapter.Speed/1000000000)
            $NetAdapterResult = $NetAdapterResult + "网卡: " + $NetAdapter.NetConnectionID + ", 速度: " + $NetAdaptersSpeed + "`r`n"
            $WarningComputer = "机器名称: $ComputerName`n"
            if($NetAdapter.Speed/1000000000 -lt $NetAdapterSpeedThreshold)
            {
                $Flag = $true
                $WarningComputer = $WarningComputer + $NetAdapter.NetConnectionID + " 网卡速度小于" + $NetAdapterSpeedThreshold + "Gbps!`r`n"
                $WarningComputersNet = $WarningComputersNet + $WarningComputer + "`r`n"
            }
        }
        }else{
        $NetAdapters = $NetAdapters | Where-Object {$_.PNPDeviceID -Like "*PCI*"}
        foreach($NetAdapter in $NetAdapters)
        {
            $NetAdaptersSpeed = "{0:0.0} Gbps" -f ($NetAdapter.Speed/1000000000)
            $NetAdapterResult = $NetAdapterResult + "网卡: " + $NetAdapter.NetConnectionID + ", 速度: " + $NetAdaptersSpeed + "`r`n"
            $WarningComputer = "机器名称: $ComputerName`n"
            if($NetAdapter.Speed/1000000000 -lt $NetAdapterSpeedThreshold)
            {
                $Flag = $true
                $WarningComputer = $WarningComputer + $NetAdapter.NetConnectionID + " 网卡速度小于" + $NetAdapterSpeedThreshold + "Gbps!`r`n"
                $WarningComputersNet = $WarningComputersNet + $WarningComputer + "`r`n"
            }
        }
        }
    }
    else{
        $NetAdapterResult = "$ComputerName 离线!`n"
    }
    $NetAdapterResults = $NetAdapterResults + $NetAdapterResult + "`r`n"
}

Write-Host "$NetAdapterResults"
Write-Host "$WarningComputersNet"

if($Flag)
{
    $MessageBox::Show("$WarningComputersNet","网卡警告")
}

$CheckData.OutString =  "------警告------`n$WarningComputersNet`n------详情------`n$NetAdapterResults" 
$CheckData.OutState = $flag