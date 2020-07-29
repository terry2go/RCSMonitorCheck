<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/04/20 修复检查本机没有显示结果的问题
== 2020/05/28 加入警告弹窗
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
        
        $LastBootTimeResult = "机器名称: $ComputerName`r`n"
        $WarningComputer = ""
        foreach($LastBootTime in $LastBootTimes)
        {
            $LastBootDate = $LastBootTime.ConvertToDateTime($LastBootTime.lastbootuptime)
            $LastBootTimeResult = $LastBootTimeResult + "上次开机时间:" + $LastBootDate + "`r`n"
            $DateResult = $LastBootDate.Date
            $BootTimeSpan = New-TimeSpan $DateResult $(Get-Date)
            Write-Host $ComputerName $BootTimeSpan.Days "天没有重启了"
            if($BootTimeSpan.Days -ge $BootTimeThreshold){
                $Flag = $true
                $WarningComputer = $WarningComputer + "$ComputerName 超过" + $BootTimeThreshold + "天没有重启了!`r`n"
                $WarningComputers = $WarningComputers + $WarningComputer + "`r`n"
            }
        }
    }
    else{
        $LastBootTimeResult = "$ComputerName 离线!`r`n"
    }
    $LastBootTimeResults = $LastBootTimeResults + $LastBootTimeResult + "`r`n"
}

Write-Host "$LastBootTimeResults"
Write-Host "$WarningComputers"

if($Flag)
{
    $MessageBox::Show("$WarningComputers","机器运行时间警告")
}

#RCS Monitor Used
$CheckData.OutString =  "------警告------`r`n$WarningComputers`r`n------详情------`r`n$LastBootTimeResults" 
$CheckData.OutState = $Flag