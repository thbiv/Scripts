<#
.SYNOPSIS
Encrypt a string then export that encrypted string to a text file.

.DESCRIPTION
This function will will encrypt a given string using the window user's key and export that encrypted string to a text file.

.PARAMETER String
The string you would like to be encrypted.

.PARAMETER Path
The path to the file that will contain the encrypted string.

.EXAMPLE
PS C:\> EXxort-EncryptedString.ps1 -String "Super Secret String" -Path "$Home\Documents\encryptedstring.txt"

.NOTES
Version: 1.0
Author: Thomas Barratt
#>
[CmdletBinding(SupportsShouldProcess=$True)]
Param (
    [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True)]
    [ValidateNotNullOrEmpty()]
    [string]$String,

    [Parameter(Mandatory=$False)]
    [ValidateScript({Test-Path -Path $_ -IsValid})]
    [ValidateScript({Test-Path -Path $(Split-Path -Path $_ -Parent)})]
    [string]$Path
)

Process {
    $Content = $String | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    If ($PSCmdlet.ShouldProcess('Encrypted String','Export')) {
        Try {
            Set-Content -Value $Content -Path $Path -ErrorAction Stop
        } Catch {
            Throw $_.Exception.Message
        }
    }
}