<#
=============================================================
== Terry Li
== 2020/04/21 初始版本
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