<#
.SYNOPSIS
Finds logged on users for a specific computer.

.DESCRIPTION
Finds logged on users for a specific computer.
Requires administrator priveleges on the computer being checked.

.PARAMETER ComputerName
The name of the computer that will be checked.

.EXAMPLE
PS C:\> Get-LoggedOnUserForComputer.ps1 -ComputerName 'computer01'

.NOTES
Version: 1.0
Author: Thomas Barratt
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True, Position=1)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName
)
$Output = @()
ForEach ($Computer in $ComputerName) {
    If ((Test-Connection -ComputerName $Computer -Count 1 -Quiet) -eq $True) {
        Write-Verbose "[$Computer] Online"
        Try {
            #$Proc = Get-WmiObject -Class Win32_Process -Computer $Computer -Filter "Name = 'explorer.exe'" -ErrorAction Stop
            $Proc = Get-CimInstance -ClassName Win32_Process -ComputerName $Computer -Filter "Name= 'explorer.exe'" -ErrorAction Stop
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Write-Warning "[$Computer] $ErrorMessage"
        }
        ForEach ($P in $Proc) {
            #$User = ($P.GetOwner()).User
            $User = Invoke-CimMethod -InputObject $P -MethodName GetOwner
            
            $Props = @{
                'ComputerName'=$Computer
                'Domain'=$($User.Domain)
                'User'=$($User.User)
            }
            $Obj = New-Object -TypeName psobject -Property $Props
            $Output += $Obj
        }
    } Else {
        Write-Verbose "[$Computer] Offline"
    }
}
Write-Output $Output