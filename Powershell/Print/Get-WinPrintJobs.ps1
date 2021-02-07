<#
Get-WinPrintJobs.ps1
Version: 1.0.0
Author: Thomas Barratt
#>
$PrintServer = $env:COMPUTERNAME
$OutputSavePath = "C:\PrinterReports"
$TranscriptSavePath = "C:\PrinterReports\Transcripts"

$FileName = "{0}-PrintJobs_{1:yyyyMMddTHHmmss}.csv" -f $PrintServer,$(Get-Date)
$TranscriptName = "Transcript_{0}-PrintJobs_{1:yyyyMMddTHHmmss}.txt" -f $PrintServer,$(Get-Date)
$OutputFile = Join-Path -Path $OutputSavePath -ChildPath $FileName
$TranscriptFile = Join-Path -Path $TranscriptSavePath -ChildPath $TranscriptName
Start-Transcript -Path $TranscriptFile -Force
$PrintOperationalLog = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration "Microsoft-Windows-PrintService/Operational"
If ($($PrintOperationalLog.IsEnabled) -eq $False) {
    Write-Warning "[$PrintServer] 'Microsoft-Windows-PrintService/Operational' log is disabled."
    $PrintOperationalLog.IsEnabled = $True
    $PrintOperationalLog.SaveChanges()
    Write-Host "[$PrintServer] 'Microsoft-Windows-PrintService/Operational' log has been enabled."
} Else { Write-Host "[$PrintServer] 'Microsoft-Windows-PrintService/Operational' log is enabled. No changes needed."}
$StopWatch = [system.diagnostics.stopwatch]::startNew()
Write-Host "Running Script on $PrintServer"
$Output = @()
$StartTime = ((get-date).AddDays(-7)).ToShortDatestring();$StartTime = (Get-Date $StartTime)
Write-Host "[$PrintServer] DateRangeStart : $StartTime"
$EndTime = ((get-date).ToShortDatestring());$EndTime = (Get-Date $EndTime)
Write-Host "[$PrintServer] DateRangeEnd : $EndTime"
Write-Host "[$PrintServer] Collecting Events"
Try {
    $PrintEvents = Get-WinEvent -FilterHashTable @{logname='Microsoft-Windows-PrintService/Operational';StartTime=$StartTime;EndTime=$EndTime}  -Oldest -ComputerName $PrintServer -ErrorAction 'Stop' | Where-Object {$_.ID -eq "307"} | Select-Object *
} Catch {
    Write-Warning "[$PrintServer] Get-WinEvent : $($Error[0])"
}
Write-Host "[$PrintServer] Adding Data to Output"
ForEach ($PrintEvent in $PrintEvents) {
    $DateCreated = $PrintEvent.TimeCreated
    $Server = $PrintServer
    $DocumentName = ($PrintEvent.Properties)[1].Value
    $User = ($PrintEvent.Properties)[2].Value
    $DisplayName = $((([adsisearcher]"samaccountname=$User").FindOne()).properties.name)
    $ComputerName = ($PrintEvent.Properties)[3].Value
    $PrinterName = ($PrintEvent.Properties)[4].Value
    $DocumentSize = $((($PrintEvent.Properties)[6].Value) / 1KB)
    $Pages = ($PrintEvent.Properties)[7].Value
    $Props = [ordered]@{
        'DateCreated'=$DateCreated
        'PrintServer'=$Server
        'Name'=$DisplayName
        'ComputerName'=$ComputerName
        'PrinterName'=$PrinterName
        'DocumentName'=$DocumentName
        'DocumentSizeInKB'=$DocumentSize
        'PageCount'=$Pages
    }
    $Obj = New-Object -TypeName PSObject -Property $Props
    $Output += $Obj
    Write-Host "[$PrintServer] Print Job for '$DisplayName' named '$DocumentName' added to output"
}
$EventCount = $Output.Count
Write-Host "[$PrintServer] Event Count : $EventCount"
Try{
    $Output | Select-Object PrintServer,DateCreated,PrinterName,Name,ComputerName,DocumentName,DocumentSizeInKB,PageCount | Export-Csv -Path $OutputFile -NoTypeInformation
} Catch {
    Write-Warning "[$PrintServer] Export-CSV : $($Error[0])"
}
Write-Host "[$PrintServer] Output Exported to CSV"
Write-Host "[$PrintServer] CSV File : $OutputFile"
$StopWatch.Stop()
Write-Host "Script Execution Time: $($StopWatch.Elapsed.Days) Days $($StopWatch.Elapsed.Hours) Hours $($StopWatch.Elapsed.Minutes) Minutes $($StopWatch.Elapsed.Seconds) Seconds $($StopWatch.Elapsed.Milliseconds) Milliseconds"
Stop-Transcript