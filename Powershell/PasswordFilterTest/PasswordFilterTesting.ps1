$Config = Import-PowershellDataFile -Path 'PasswordFilterTesting.config.psd1'

$Servers = @()
If($($Config.IncludeAllDomainControllers) -eq $True) {
    $Servers += Get-ADDomainController -filter {isReadOnly -eq $False} | Select-Object -ExpandProperty HostName
}
$Servers += $($Config.IncludeComputerNames)

$PesterParams = @{
    Script = @{Path = $PSScriptRoot; Parameters = @{SourceDLLPath = $($Config.SourceDLLPath); Servers = @()}}
    PassThru = $True
}

$Results = Invoke-Pester @PesterParams

$PreContent = @()
$PreContent += "Total Count: $($Results.TotalCount)"
$PreContent += "Passed Count: $($Results.PassedCount)"
$PreContent += "Failed Count: $($Results.FailedCount)"
$PreContent += "Duration: $($Results.Time)"

$MailSubject = "Password Filter Testing Results [{0} Passed] [{1} Failed]" -f $($Results.PassedCount),$($Results.FailedCount)
$Header = @"
<style>
TABLE{border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH{border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD{border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
TR:nth-child(odd){background:#FFFFFF;}
TR:nth-child(even){background:#808080;}
</style>
"@
$MailProps = @{
    'SmtpServer' = $($Config.MailSettings.SMTPServer)
    'Port' = $($Config.MailSettings.Port)
    'To' = $($Config.MailSettings.To)
    'From' = $($Config.MailSettings.From)
    'Subject' = $MailSubject
    'Body' = $($Results.TestResult | ConvertTo-Html -Property Describe,Context,Name,Result,Time,FailureMessage,StackTrace,ErrorRecord -Head $Header -PreContent $($PreContent -join '<BR>') | Out-String)
    'BodyAsHtml' = $True
}
If ($($Config.SendMail) -eq $True) {
    Send-MailMessage @MailProps
}