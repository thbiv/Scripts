<#
.SYNOPSIS
Deletes files and folders when you cannot delete them normally because of the 260 character path limit.

.DESCRIPTION
Uses ROBOCOPY to delete files and folders by mirroring an empty folder to the folder you wish to delete.

.PARAMETER Path
Path to the folder that will have all its files and folders deleted.

.EXAMPLE
PS C:\> Remove-FilesOverPAthLimit -Path <path to folder>
    
.INPUTS
None

.OUTPUTS
None

.NOTES
Version: 1.0.0
Author: Thomas Barratt
#>
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path -Path $_})]
    [string]$Path
)

[string]$Source = $(Join-Path -Path $env:LOCALAPPDATA -ChildPath 'EmptyFolder')

If ((Test-Path -Path $Source) -eq $False) {
    Write-Verbose "Source Does Not Exist :: Creating Folder"
    Try {
        New-Item -Path $Source -ItemType Directory -Force -WhatIf:$WhatIfPreference | Out-Null
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Write-Warning $ErrorMessage
    }
} Else {
    Try {
        Remove-Item -Path "$Source\*.*" -Force -Confirm:$false -WhatIf:$WhatIfPreference | Out-Null
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Write-Warning $ErrorMessage
    }
}
$RobocopyParams = "/MIR /NP /NJH /NJS /NC"
[string]$ArgumentList = ('"{0}" "{1}" {2}' -f $Source, $Path, $RobocopyParams)
Try {
    If ($PSCmdlet.ShouldProcess($Path,'Delete')) {
        Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -NoNewWindow -Wait -ErrorAction Stop | Out-Null
    }
} Catch {
    $ErrorMessage = $_.Exception.Message
    Write-Warning $ErrorMessage
}