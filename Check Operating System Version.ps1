<#
=============================================================
== Terry Li
== 2020/04/21 初始版本
== 2020/07/14 加入忽略运行过程中的错误
== 2020/07/27 将远程计算机的用户名以及密码单独写到ini配置文件，便于集中修改
=============================================================
#>
$ErrorActionPreference= "silentlycontinue"
$ComputerList = @()
$Path = 'D:\Bat\RCSMonitor'

$settingsKeys = @{
    UserName = "^\s*UserName\s*$";
    Password = "^\s*Password\s*$";
    ComputerName = "^\s*ComputerName\s*$";
}

Get-Content $Path\config.ini | Foreach-Object {  
    $var = $_.Split('=')
    $settingsKeys.Keys |% {
        if ($var[0] -match $settingsKeys.Item($_))
        {
            if ($_ -eq 'ComputerName')
            {
                $ComputerList += $var[1].Trim()
            }
            elseif ($_ -eq 'UserName')
            {
                $UserName = $var[1].Trim()
            }
            elseif ($_ -eq 'Password')
            {
                $Password = $var[1].Trim()
            }
            else
            {
                New-Variable -Name $_ -Value $var[1].Trim() -ErrorAction silentlycontinue
            }
        }
    }
}

$PasswordNew = ConvertTo-SecureString $Password -AsPlainText -Force;
$Cred = New-Object System.Management.Automation.PSCredential($UserName,$PasswordNew)

$Flag = $false
$BuildVersionResults = ""

foreach ($ComputerName in $ComputerList){
    $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
    $PingResult = Get-WmiObject -query $PingQuery
    if($PingResult.ProtocolAddress){

    if($PingResult.__SERVER -like $ComputerName){
        $OperatingSystems = Get-WmiObject -class Win32_OperatingSystem -computername $ComputerName
    }
    else{
        $OperatingSystems = Get-WmiObject -class Win32_OperatingSystem -Credential $Cred -computername $ComputerName
    }
        
        $BuildVersionResult = "机器名称: $ComputerName`n"
        foreach($OperatingSystem in $OperatingSystems)
        {
            $OperatingSystemCaption = $OperatingSystem.Caption
            $OperatingSystemBuildNumber = $OperatingSystem.BuildNumber
            $BuildVersionResult = $BuildVersionResult + "操作系统: " + $OperatingSystemCaption + ", 版本号: " + $OperatingSystemBuildNumber +"`n"

        }
    }
    else{
        $BuildVersionResult = "$ComputerName 离线!`n"
    }
    $BuildVersionResults = $BuildVersionResults + $BuildVersionResult + "`n"
}

Write-Host "$BuildVersionResults"

#RCS Monitor Used
$CheckData.OutString =  "`n------详情------`n$BuildVersionResults" 
$CheckData.OutState = $Flag