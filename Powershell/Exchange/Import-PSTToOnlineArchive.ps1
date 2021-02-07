<#
.SYNOPSIS
    Script for assisting in the process of Migrating PST files to Online Archives in Microsoft 365.
.DESCRIPTION
    This script will upload PST files from a folder to Azure Storage, then creates the mapping file for those files
    to use when migrating the data. The TargetRootFolder for each file will be the file name itself. Please make sure
    the file names are descriptive.
.PARAMETER Path
    The folder that contains the PST files to upload. Must not contain spaces.
.PARAMETER SASURL
    The SAS URL that you get from the Import Wizard in Microsoft 365
.PARAMETER User
    The logon name for the user in Microsoft 365.
.PARAMETER OutputPath
    The path to the exported mapping file.
    Defaults to the Desktop of the user that runs this script.
#>
[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$True)]
    [string]$Path,

    [Parameter(Mandatory=$True)]
    [string]$SASURL,

    [Parameter(Mandatory=$True)]
    [string]$User,

    [Parameter(Mandatory=$False)]
    [string]$OutputPath = $(Join-Path -Path $Home -ChildPath 'Desktop')
)

$TimeStamp = $(Get-Date)

$LogName = "AZCopy-Log_{0:yyyMMdd}.log" -f $TimeStamp
$Log = Join-Path -Path $OutputPath -ChildPath $LogName
Write-Verbose -Message "[$(($User.Split('@'))[0])] AzCopy Log Path: $Log"

Write-Verbose -Message "[$(($User.Split('@'))[0])]Starting AzCopy.exe...Uploading PST Files to Azure Storage"
$AZCopy = "C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy\AzCopy.exe"
$ArgumentList = "/Source:{0} /Dest:{1} /V:{2} /Y" -f $Path,$SASURL,$Log
If ($PSCmdlet.ShouldProcess("Uploading PST files from: $Source")) {
    Start-Process -FilePath $AZCopy -ArgumentList $ArgumentList -NoNewWindow -Wait
    Write-Verbose -Message "[$(($User.Split('@'))[0])] AzCopy Process Complete."
}

$MappingFileName = 'ImportMappingFile_{0}_{1:yyyMMdd}.csv' -f $(($User.Split('@'))[0]),$TimeStamp
$MappingFile = Join-Path -Path $OutputPath -ChildPath $MappingFileName
$PSTList = Get-ChildItem -Path $Path | Select-Object -ExpandProperty Name
$Output = @()
ForEach ($PST in $PSTList) {
    $Props = [ordered]@{
        'WorkLoad'='Exchange'
        'FilePath'=$Null
        'Name'=$PST
        'Mailbox'=$User
        'IsArchive'='TRUE'
        'TargetRootFolder'="/$(($PST.split('.'))[0])"
        'SPFileContainer'=$Null
        'SPManifestContainer'=$Null
        'SPSiteUrl'=$Null
    }
    $Obj = New-Object -TypeName PSObject -Property $Props
    $Output += $Obj
    Write-Verbose -Message "[$(($User.Split('@'))[0])] PST File Added: $PST"
}
If ($PSCmdlet.ShouldProcess("Exporting Mapping File: $MappingFile")) {
    $Output | Export-Csv -Path $MappingFile -NoTypeInformation -Force
}