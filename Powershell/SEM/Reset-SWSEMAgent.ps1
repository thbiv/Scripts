<#
.SYNOPSIS
    Resets the Solarwinds Security Event Manager Agent so that it will start talking to the server again.
.DESCRIPTION
    This function reset the SEM Agent remotely by stopping the two services, deleting the spop folder, then starting the services again.
.PARAMETER ComputerName
    The name of the computer you wish to remotely reset the SEM agent on.
    Cannot be used with the Path parameter.
.PARAMETER Path
    The path to a text file with a list of computernames, one per line, you wish to remotely reset the SEM agent on.
    Cannot be used with the ComputerName parameter.
.EXAMPLE
    PS C:\> Reset-SWSEMAgent -ComputerName server01,server02

    This example uses the ComputerName parameter to perform the reset on server01 and server02.
.EXAMPLE
    PS C:\> Reset-SWSEMAgent -Path Names.txt

    This example uses the path parameter to import a list of names from a text file named Names.txt.
.EXAMPLE
    PS C:\> Reset-SWSEMAgent

    This example does not use any parameters. This tells the function that the repair is to be run locally.
.INPUTS
    String
.OUTPUTS
#>

[CmdletBinding(SupportsShouldProcess=$True,DefaultParameterSetName='ComputerName')]
Param (
    [Parameter(Mandatory=$False,Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,ParameterSetName='ComputerName')]
    [Alias('HostName','Host')]
    [string[]]$ComputerName,

    [Parameter(Mandatory=$True,ParameterSetName='Path')]
    [ValidateScript({Test-Path $_})]
    [string]$Path
)

Begin {
    Function ExecuteRemote {
        Param (
            [string]$Name
        )
        Try {
            $Session = New-PSSession -ComputerName $Name -ErrorAction STOP
            Invoke-Command -Session $Session -ErrorAction Stop -ScriptBlock {
                $Services = @('Contego_Spop','USB-Defender')
                $Paths = @('C:\Windows\sysWOW64\contegoSPOP\spop','C:\Windows\system32\ContegoSPOP\spop')
                ForEach ($Service in $Services) {
                    $S = Get-Service -Name $Service
                    If ($S) {$S | Stop-Service -Force}
                    Write-Verbose "[$Name] Service Stopped: $Service"
                }
                ForEach ($Item in $Paths) {
                    If (Test-Path -Path $Item) {
                        Remove-Item -Path $Item -Force -Recurse
                        Write-Verbose "[$Name] Folder Deleted: $Item"
                    }
                }
                ForEach ($Service in $Services) {
                    $S = Get-Service -Name $Service
                    If ($S) {$S | Start-Service}
                    Write-Verbose "[$Name] Service Started: $Service"
                }
            }
            Remove-PSSession -Session $Session
            Write-Verbose "[$Name] Complete"
        } Catch {
            $ErrorMessage = $_.exception.message
            Write-Warning "[$Name] ERROR: $ErrorMessage"
        }
    }
    Function ExecuteLocal {
        $Services = @('Contego_Spop','USB-Defender')
        $Paths = @('C:\Windows\sysWOW64\contegoSPOP\spop','C:\Windows\system32\ContegoSPOP\spop')
        ForEach ($Service in $Services) {
            $S = Get-Service -Name $Service
            If ($S) {$S | Stop-Service -Force}
            Write-Verbose "[$Name] Service Stopped: $Service"
        }
        ForEach ($Item in $Paths) {
            If (Test-Path -Path $Item) {
                Remove-Item -Path $Item -Force -Recurse
                Write-Verbose "[$Name] Folder Deleted: $Item"
            }
        }
        ForEach ($Service in $Services) {
            $S = Get-Service -Name $Service
            If ($S) {$S | Start-Service}
            Write-Verbose "[$Name] Service Started: $Service"
        }
    }
}

Process {
    If ($PSCmdlet.ParameterSetName -eq 'ComputerName') {
        If ($ComputerName) {
            ForEach ($Name in $ComputerName) {
                If (Test-Connection -ComputerName $Name -Count 1 -Quiet) {
                    Write-Verbose "[$Name] Is Online"
                    If ($PSCmdlet.ShouldProcess("$Name","Repair SEM Agent")) {
                        ExecuteRemote -Name $Name
                    }

                } Else {
                    Write-warning "[$Name] Is Offline"
                }
            }
        } Else {
            If ($PSCmdlet.ShouldProcess("$Name","Repair SEM Agent")) {
                ExecuteLocal
            }
        }
    }
    If ($PSCmdlet.ParameterSetName -eq 'Path') {
        $Nodes = Get-Content -Path $Path
        ForEach ($Node in $Nodes) {
            If (Test-Connection -ComputerName $Node -Count 1 -Quiet) {
                Write-Verbose "[$Node] IsOnline"
                If ($PSCmdlet.ShouldProcess("$Node","Repair SEM Agent")) {
                    ExecuteRemote -Name $Node
                }
            } Else {
                Write-warning "[$Node] IsOffline"
            }
        }
    }
}