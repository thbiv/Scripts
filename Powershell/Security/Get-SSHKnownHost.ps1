<#
.Synopsis
Parse ssh known hosts
.Description
This command will parse the ssh known_hosts file and write custom objects to
the pipeline. If HashKnownHosts is set to yes in ssh_config, you may not get
meaningful results.
.Parameter Hostname
Specify a hostname. Wildcards are permitted. Leave blank to see all known
ssh hosts.
.Example
PS C:\> Get-sshKnownHost
Hostname   : localhost
Port       :
Address    :
Keytype    : ecdsa-sha2-nistp256
Thumbprint : AAAAE2VjZHNhLXNoYTItbmlzdHAyNTYAAAAIbmlzdHAyNTYAAABBBGCbSDjrVN
                B/oGITy7qKovJam+k2HKYCtJzyiYTjevW/mYIF5umy/1eOG7Nb2AtNgpI0p6ah
                ZtChttdT/hcZfAU=
Hostname   : 192.168.3.199
Port       :
Address    :
Keytype    : ecdsa-sha2-nistp256
Thumbprint : AAAAE2VjZHNhLXNoYTItbmlzdHAyNTYAAAAIbmlzdHAyNTYAAABBBOZuPhlOpx
                sbkWNT2YKr4OgjSz05B0CYfHbmPlFeshOVo3r5GaXDZ+N5aRfbbPf6VYqMK/XP
                gO7VQ2rNcHdAxOw=
...
.Example
PS C:\> Get-KnownHost -hostname srv*
Hostname   : srv1
Port       :
Address    : 192.168.3.50
Keytype    : ecdsa-sha2-nistp256
Thumbprint : AAAAE2VjZHNhLXNoYTItbmlzdHAyNTYAAAAIbmlzdHAyNTYAAABBBP6unwvoWX
                AbQ2ymqO3/TB2zSBayXP1ke2J+YxxOe57WoJ9ZEWdDyNdXwjYPzO139eVa8gFz
                gPSV4DgDm/hLfYM=
Hostname   : srv2
Port       :
Address    : 192.168.3.51
Keytype    : ecdsa-sha2-nistp256
Thumbprint : AAAAE2VjZHNhLXNoYTItbmlzdHAyNTYAAAAIbmlzdHAyNTYAAABBBKcUMuWoDN
                oM0RojPT+p/z8F2N7pPychlc49oTvCsGH5urCaTu6R4Fu9tctxLvuPKRIpl08+
                DRTnXOGI3m/wDbw=
.Notes
Learn more about PowerShell: http://jdhitsolutions.com/blog/essential-powershell-resources/
.Link
ssh-keygen
.Inputs
none
.Outputs
sshKnownHost
#>
[cmdletbinding()]
[alias("gkh")]
[Outputtype("sshKnownHost")]
Param(
    [Parameter(Position = 0, HelpMessage = "Specify a hostname. Wildcards are permitted. Leave blank to see all known hosts.")]
    [string]$Hostname
)

Write-Verbose "Starting $($MyInvocation.MyCommand)"
#define a list to hold host data
$data = [System.Collections.Generic.list[object]]::New()
[regex]$rx = "(?<host>^\S+?)((?=:))?((?<=:)(?<port>\d+))?(,(?<address>\S+))?\s(?<type>[\w-]+)\s(?<thumbprint>.*)"
$known = "~\.ssh\known_hosts"
Write-Verbose "Testing for $known"
if (Test-Path $known) {
    $content = (Get-Content -Path $known) -split "`n"
    Write-Verbose "Found $($content.count) entries"

    #process all entries even if searching for a single hostname because there
    #might be multiple entries
    foreach ($entry in $content) {
        $matched = $rx.Match($entry)
        $sshHost = $matched.groups["host"].value -replace ":$|\[|\]", ""
        $IP = $matched.groups["address"].value -replace "\[|\]", ""
        Write-Verbose "Processing $sshHost"
        #regex named captures are case-sensitive
        #I haven't perfected the regex capture so I'll manually trim and trailing : in the hostname capture
        $obj = [pscustomobject]@{
            PSTypeName = "sshKnownHost"
            Hostname   = $sshHost
            Port       = $matched.groups["port"].value
            Address    = $IP
            Keytype    = $matched.groups["type"].value
            Thumbprint = $matched.groups["thumbprint"].value
        }            #add each new object to the list
        $data.Add($obj)
    } #foreach entry
}
else {
    Write-Warning "Can't find $known. Is the ssh client installed?"
}

if ($Hostname -AND $data.count -gt 0) {
    Write-Verbose "Searching for $hostname"
    $data.where({$_.hostname -like $hostname})
}
elseif ($data.count -gt 0) {
    $data
}
Write-Verbose "Ending $($MyInvocation.MyCommand)"