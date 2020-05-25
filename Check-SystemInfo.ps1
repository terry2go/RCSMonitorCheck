function Check-SystemInfo {
    [CmdletBinding(SupportsShouldProcess=$false)]
    Param(
        [Parameter(Position=0, Mandatory=$false)] [PSCredential]$Credential,
        [Parameter(Position=1, Mandatory=$true)] [string]$Computername
        )
    Process {
        $UserName = "RCS-USER"
        $Password = ConvertTo-SecureString "RCS" -AsPlainText -Force;
        $Cred = New-Object System.Management.Automation.PSCredential($UserName,$Password)
        $PingQuery = "select * from win32_pingstatus where address = '$ComputerName'"
        $PingResult = Get-WmiObject -query $PingQuery
        if($PingResult.ProtocolAddress){
            if($PingResult.__SERVER -like $ComputerName){
                $SystemInfo_Bios = Get-WmiObject -Class win32_bios -ComputerName $ComputerName
                $SystemInfo_OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
                $SystemInfo_Memory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $ComputerName
                $SystemInfo_Processor = Get-WmiObject -Class Win32_Processor -ComputerName $ComputerName
                $SystemInfo_Disk = Get-WmiObject -class Win32_DiskDrive -ComputerName $ComputerName
                $SystemInfo_Video = Get-WmiObject -Class CIM_VideoController -ComputerName $ComputerName
                }
            else{
                $SystemInfo_Bios = Get-WmiObject -Class win32_bios -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_OS = Get-WmiObject -Class Win32_OperatingSystem -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Memory = Get-WmiObject -Class Win32_PhysicalMemory -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Processor = Get-WmiObject -Class Win32_Processor -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Disk = Get-WmiObject -class Win32_DiskDrive -Credential $Cred -ComputerName $ComputerName
                $SystemInfo_Video = Get-WmiObject -Class CIM_VideoController -Credential $Cred -ComputerName $ComputerName
                }
            Write-Host "== "$ComputerName" 配置`r`n"
            $SystemInfo_Bios_Num = $SystemInfo_Bios.serialnumber
            $SystemInfo_Bios_Num_Result = "机器序列号: " + $SystemInfo_Bios_Num + "`r`n"
            Write-Host $SystemInfo_Bios_Num_Result
            $SystemInfo_OS_Caption = $SystemInfo_OS.caption
            $SystemInfo_OS_Build = $SystemInfo_OS.BuildNumber
            $SystemInfo_OS_Result = "操作系统: " + $SystemInfo_OS_Caption + " Build " + $SystemInfo_OS_Build + "`r`n"
            Write-Host $SystemInfo_OS_Result
            Write-Host "`r`n== 内存信息`r`n"
            $SystemInfo_Memory | Select-Object Manufacturer,SerialNumber,Capacity | ForEach-Object{
                $SystemInfo_Memory_Result = "内存品牌: " + $_.Manufacturer + ", 序列号: " + $_.SerialNumber + ", 容量: " + $_.Capacity/1GB + "G"
                $SystemInfo_Memory_Result2 = $SystemInfo_Memory_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Memory_Result2
                $SystemInfo_Memory_Results = $SystemInfo_Memory_Results + $SystemInfo_Memory_Result2 + "`r`n"
                }
            Write-Host "`r`n== CPU信息`r`n"
            $SystemInfo_Processor | Select-Object Manufacturer,Name,NumberOfCores,NumberOfLogicalProcessors | ForEach-Object{
                $SystemInfo_Processor_Result = "CPU: " + $_.Name + ", " + $_.NumberOfCores + "核, " + $_.NumberOfLogicalProcessors + "线程"
                $SystemInfo_Processor_Result2 = $SystemInfo_Processor_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Processor_Result2
                $SystemInfo_Processor_Results = $SystemInfo_Processor_Results + $SystemInfo_Processor_Result2 + "`r`n"
                }
            Write-Host "`r`n== 硬盘信息`r`n"
            $SystemInfo_Disk | Select-Object Model,SerialNumber,Size | Foreach-Object{
                $SystemInfo_Disk_Result = "硬盘: " + $_.model + ", 序列号: " + $_.SerialNumber + ", 容量: " + [math]::Round($_.size/1GB) + "GB"
                $SystemInfo_Disk_Result2 = $SystemInfo_Disk_Result -replace '\s{2,}', ' '
                Write-Host $SystemInfo_Disk_Result2
                $SystemInfo_Disk_Results = $SystemInfo_Disk_Results + $SystemInfo_Disk_Result2 + "`r`n"
                }
            Write-Host "`r`n== 显卡信息`r`n"
            $SystemInfo_Video | select Caption | ForEach-Object{
                $SystemInfo_Video_Result = "显卡: " + $_.Caption
                Write-Host $SystemInfo_Video_Result
                $SystemInfo_Video_Results = $SystemInfo_Video_Results + $SystemInfo_Video_Result + "`r`n"
                }
            $SystemInfo_Result = "== " + $ComputerName + " 配置`r`n" + $SystemInfo_Bios_Num_Result + "`r`n" + $SystemInfo_OS_Result + "`r`n" + $SystemInfo_Memory_Results + "`r`n"  + $SystemInfo_Processor_Results+ "`r`n"   + $SystemInfo_Video_Results
            $SystemInfo_Result
        }
    }
}

Check-SystemInfo -Computername RCSWH-SVR01