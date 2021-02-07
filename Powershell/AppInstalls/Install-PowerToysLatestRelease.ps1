<#
.SYNOPSIS
Downloads and starts the install for Microsoft Power Toys
.DESCRIPTION
Downloads and starts the install for Microsoft Power Toys
.EXAMPLE
PS C:\> Install-PowerToysLatestRelease.ps1
.INPUTS
None
.OUTPUTS
None
.NOTES
Version: 1.0.0
Author: Thomas Barratt
#>

$Params = @{
    'Uri' = 'https://api.github.com/repos/Microsoft/PowerToys/releases/latest'
    'Headers' = @{"Accept"="application/json"}
    'Method' = 'Get'
    'UseBasicParsing' = $True
}
$Response = Invoke-RestMethod @Params
$Assets = $Response.assets | Where-Object {$_.name -like '*.exe'}
Invoke-WebRequest -Uri $($Assets.browser_download_url) -OutFile "$env:temp\$($Assets.name)"
& "$env:temp\$($Assets.name)"