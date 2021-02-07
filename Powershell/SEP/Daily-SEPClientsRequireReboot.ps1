#Requires -Module SEPM

$DateTimeStamp = $(Get-Date)
$LogFile = "$PSScriptRoot\Logs\SEPClientsRequireReboot.log" -f $DateTimeStamp
$ServerName = '' # Name of the Symantec Endpoint Protection Manager server
$CredentialFile = "$Home\Documents\SEPMCreds.xml"

$PreContent = @()
$PreContent += "DateTime: $DateTimeStamp"
$PreContent += "SEPM Server: $ServerName"

If (!(Test-Path -Path $LogFile)) {
    New-Item -Path $LogFile -ItemType File -Force | Out-Null
}
$Msg = "ScriptStart: $DateTimeStamp"
Add-Content -Path $LogFile -Value $Msg
Write-Information -InformationAction Continue -MessageData $Msg

If (!(Test-Path -Path $CredentialFile)) {
    $Msg = "Credential File not Found: $CredentialFile"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
    Exit
}

$Cred = Import-Clixml -Path $CredentialFile
Try {
    $Token = Get-SEPMAccessToken -ComputerName $ServerName -Credential $Cred -ErrorAction Stop
    $Msg = "[Get-SEPMAccessToken] Token Successfully Retrieved for: $($Cred.UserName)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
} Catch {
    $Msg = "[Get-SEPMAccessToken] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
}

Try {
    $Clients = Get-SEPMClient -ComputerName $ServerName -Token $($Token.Token) -ErrorAction Stop
    $Msg = "[Get-SEPMClient] Retrieved all clients from $ServerName : Total: $($Clients.Count)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
} Catch {
    $Msg = "[Get-SEPMClient] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
}

$PreContent += "Total Clients: $($Clients.Count)"

Try {
    $ClientNeedReboot = @()
    ForEach ($Client in $Clients) {
        If ($client.RebootRequired -eq $True) {
            $Props = [ordered]@{
                'ComputerName' = $($Client.ComputerName)
                'RebootRequired' = $($Client.RebootRequired)
                'RebootReason' = $($Client.RebootReason)
                'Group' = $($Client.Group)
            }
            $ClientNeedReboot += $(New-Object -TypeName PSObject -Property $Props)
            $Msg = "[$($Client.ComputerName)] Reboot Required"
            Add-Content -Path $LogFile -Value $Msg
            Write-Information -InformationAction Continue -MessageData $Msg
        }
    }
    $Msg = "[SEPClientsRequireReboot] Calculated clients that need a reboot : Total: $($ClientNeedReboot.Count)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
} Catch {
    $Msg = "[SEPClientsRequireReboot] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
    Exit
}

$PreContent += "Number of Clients that Require a Reboot: $($ClientNeedReboot.Count)"
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
TR:nth-child(even) {background: #CCC}
TR:nth-child(odd) {background: #FFF}
</style>
"@
Try {
    $MailProps = @{
        'SmtpServer' = ""
        'Port' = "25"
        'To' = ""
        'From' = ""
        'Subject' = "Symantec Endpoint Protection: Clients that Require a Restart: $($ClientNeedReboot.count)"
        'Body' = $($ClientNeedReboot | ConvertTo-Html -PreContent $($PreContent -Join '<BR>') -Head $Header | Out-String)
        'BodyAsHtml' = $True
    }
    Send-MailMessage @MailProps
} Catch {
    $Msg = "[Send-MailMessage] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
}