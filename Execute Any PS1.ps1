$1 = Get-Content "D:\Bat\RCSMonitor\RCSMonitorCheck\Check CPU Usage.ps1" -Raw
$ScriptBlock = [System.Management.Automation.ScriptBlock]::Create($1)
&$ScriptBlock