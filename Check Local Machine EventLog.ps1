<#
=============================================================
== Terry Li
== 2020/04/17 初始版本
== 2020/07/14 加入忽略运行过程中的错误
=============================================================
#>
$ErrorActionPreference= "silentlycontinue"
$StartTime = [datetime]::today
$EndTime = [datetime]::now
$Flag = $false

$EventFilter = @{Logname='System','Application'
                 Level=2,3
                 StartTime=$StartTime
                 EndTime=$EndTime
                 ProviderName='e1dexpress','*IAStor*','*SQL*'
                 }      

$events = Get-WinEvent -FilterHashtable $EventFilter -MaxEvents 5
$results = ""
foreach($event in $events)
{
    $EventTime = $event.TimeCreated
    $EventName = $event.ProviderName
    $EventLevel = $event.LevelDisplayName
    $EventMessage = $event.Message
    $result = "$EventTime $EventName $EventLevel`n"
    $results = $results + $result + "`n"
    $EventName.GetType()
    if($EventName){
        $Flag = $true
    }
}
Write-Host $results
Write-Host $flag

$CheckData.OutString =  "------详情------`n$results" 
$CheckData.OutState = $Flag