$SourceFile = 'nDepthData.csv'
$DestinationPath = '\\sfhouitrpt01\e$\Reports\ITReports\SWReports\Network\RemoteAccess'

$ImportedData = Import-Csv -Path $SourceFile
$Data = @()
$ObjectCount = 0
ForEach ($Event in $ImportedData) {
    $ObjectCount++
    $Date = ([datetime]($Event.DetectionTime).TrimEnd(":000")).ToString("yyyyMMdd")
    $Event | Add-Member -MemberType NoteProperty -Name 'Date' -Value $Date
    $Data += $Event
}
$UniqueList = $Data | Select-Object -ExpandProperty Date -Unique

ForEach ($Item in $UniqueList) {
    $FileName = "VPN_{0}.csv" -f $Item
    $Data | Where-Object {$_.Date -eq $Item} |
        Select-Object 'Event Name',EventInfo,DetectionIP,DetectionTime,InsertionIP,Manager,InsertionTime,Severity,ToolAlias,ProviderSID,ExtraneousInfo,DestinationDomain,SourceAccount,IsThreat,DestinationAccount,PrivilegesExercised,SourceMachine |
        Export-Csv -Path $(Join-Path -Path $DestinationPath -ChildPath $FileName) -NoTypeInformation -Force
}
Remove-Item -Path $SourceFile -Force