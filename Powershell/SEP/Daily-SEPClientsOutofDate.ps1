#Requires -Module SEPM

$DateTimeStamp = $(Get-Date)
$LogFile = "$PSScriptRoot\Logs\OutofDateSEPClients.log"
#$OutputFile = "$PSScriptRoot\Output\OutofDateSEPClients_{0:yyyyMMddTHHmmss}.csv" -f $DateTimeStamp
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
    $SEPMVirusDefVersion = (Get-SEPMCurrentVirusDef -ComputerName $ServerName -Token $($Token.Token)-ErrorAction Stop).PublishedBySEPM
    $SEPMVirusDefDate = [datetime]($SEPMVirusDefVersion -split " ")[0]
    $Msg = "[Get-SEPMCurrentVirusDef] Latest Virus Def Date: $SEPMVirusDefDate"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
} Catch {
    $Msg = "[Get-SEPMCurrentVirusDef] Error: $($_.Exception.Message)"
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
    $OutofDateClients = @()
    ForEach ($Client in $Clients) {
        If ($($Client.VirusDefDate) -ne 'Unknown') {
            $VirusDefAge = (New-TimeSpan -Start $($Client.VirusDefDate) -End $SEPMVirusDefDate -ErrorAction Stop).TotalDays
            If ($VirusDefAge -gt 10) {
                $Props = [ordered]@{
                    'Id' = $($Client.Id)
                    'ComputerName' = $($Client.ComputerName)
                    'Group' = $($Client.Group)
                    'VirusDefVersion' = $($Client.VirusDefVersion)
                    'VirusDefAge' = $VirusDefAge
                }
                $Obj = New-Object -TypeName PSObject -Property $Props
                $OutofDateClients += $Obj
            }
        } Else {
            $Msg = "[OutofDateClients][$($Client.ComputerName)] Virus Def Version is 'Unknown'"
            Add-Content -Path $LogFile -Value $Msg
            Write-Warning $Msg
        }
    }
    $Msg = "[OutofDateClients] Calculated Out of Date Client : Total: $($OutofDateClients.Count)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
} Catch {
    $Msg = "[OutofDateClients][$($Client.ComputerName)] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
}

$PreContent += "Out of Date Clients: $($OutofDateClients.Count)"

$Output = @()
$IDs = @()
ForEach ($OutofDateClient in $OutofDateClients) {
    $Obj = $OutofDateClient
    $IsOnline = $(Test-Connection -ComputerName $($OutofDateClient.ComputerName) -Count 1 -Quiet)
    $Obj | Add-Member -MemberType NoteProperty -Name 'IsOnline' -Value $IsOnline
    If ($IsOnline -eq $True) {
        $Msg = "[$($OutofDateClient.ComputerName)] Is Online"
        Add-Content -Path $LogFile -Value $Msg
        Write-Information -InformationAction Continue -MessageData $Msg

        Try {
            $ServiceObject = Get-Service -ComputerName $($OutofDateClient.ComputerName) -Name SepMasterService
            $Obj | Add-Member -MemberType NoteProperty -Name 'ServiceStatus' -Value $($ServiceObject.Status)
        } Catch {
            $Msg = "[$($OutofDateClient.ComputerName)] Error retrieving service status: $($_.Exception.Message)"
            Add-Content -Path $LogFile -Value $Msg
            Write-Warning -Message $Msg
        }
        If ($($ServiceObject.Status) -eq 'Stopped') {
            $Msg = "[$($OutofDateClient.ComputerName)] SepMasterService is Stopped"
            Add-Content -Path $LogFile -Value $Msg
            Write-Information -InformationAction Continue -MessageData $Msg
            Try {
                $ServiceObject | Start-Service -ErrorAction Stop
                $Msg = "[$($OutofDateClient.ComputerName)] SepMasterService Started Successfully"
                Add-Content -Path $LogFile -Value $Msg
                Write-Information -InformationAction Continue -MessageData $Msg
                $Obj | Add-Member -MemberType NoteProperty -Name 'UpdateAllowed' -Value $True
                $IDs += $($OutofDateClient.Id)
            } Catch {
                $Msg = "[$($OutofDateClient.ComputerName)] Error Starting SepMasterService: $($_.Exception.Message)"
                Add-Content -Path $LogFile -Value $Msg
                Write-Warning -Message $Msg
                $Obj | Add-Member -MemberType NoteProperty -Name 'UpdateAllowed' -Value $False
            }
        } ElseIf ($($ServiceObject.Status) -eq 'Running') {
            $Msg = "[$($OutofDateClient.ComputerName)] SepMasterService is Running"
            Add-Content -Path $LogFile -Value $Msg
            Write-Information -InformationAction Continue -MessageData $Msg
            $Obj | Add-Member -MemberType NoteProperty -Name 'UpdateAllowed' -Value $True
            $IDs += $($OutofDateClient.Id)
        } Else {
            $Msg = "[$($OutofDateClient.ComputerName)] SepMasterService is $($ServiceObject.Status)"
            Add-Content -Path $LogFile -Value $Msg
            Write-Information -InformationAction Continue -MessageData $Msg
        }
    } Else {
        $Msg = "[$($OutofDateClient.ComputerName)] Is Offline"
        Add-Content -Path $LogFile -Value $Msg
        Write-Information -InformationAction Continue -MessageData $Msg
        $Obj | Add-Member -MemberType NoteProperty -Name 'ServiceStatus' -Value 'N/A'
        $Obj | Add-Member -MemberType NoteProperty -Name 'UpdateAllowed' -Value $False
    }
    $Output += $Obj
}

$Msg = "Clients for Content Update Request: $($IDs.Count)"
Add-Content -Path $LogFile -Value $Msg
Write-Information -InformationAction Continue -MessageData $Msg

$PreContent += "Clients Requested for Content Update: $($IDs.Count)"
If ($($IDs.Count) -ne 0) {
    Try {
        $CommandResult = Update-SEPMClientContent -ComputerName $ServerName -ComputerId $($IDs -join ',') -Token $($Token.Token) -ErrorAction Stop
        $CommandResult
        $Msg = "SEPM CommandId: $($CommandResult.CommandId)"
        Add-Content -Path $LogFile -Value $Msg
    } Catch {
        $Msg = "[Update-SEPMClientContent] Error: $($_.Exception.Message)"
        Add-Content -Path $LogFile -Value $Msg
        Write-Warning $Msg
    }
    $PreContent += "Content Update Command ID: $($CommandResult.CommandId)"
} Else {
    $Msg = "[OutofDateClients] 0 clients available for content update request"
    Add-Content -Path $LogFile -Value $Msg
    Write-Information -InformationAction Continue -MessageData $Msg
}
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
TR:nth-child(even) {background: #CCC}
TR:nth-child(odd) {background: #FFF}
</style>
"@
$MailBody = $Output | ConvertTo-Html -PreContent $($PreContent -Join '<BR>') -Head $Header | Out-String
Try {
    $MailProps = @{
        'SmtpServer' = ""
        'Port' = 25
        'To' = ""
        'From' = ""
        'Subject' = "Symantec Endpoint Protection: Clients with Out of Date Virus Definitions: $($IDs.Count)/$($OutofDateClients.Count)"
        'Body' = $MailBody
        'BodyAsHtml' = $True
    }
    Send-MailMessage @MailProps -ErrorAction Stop
} Catch {
    $Msg = "[Send-MailMessage] Error: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $Msg
    Write-Warning $Msg
}