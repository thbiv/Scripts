[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True,Position=0)]
    [ValidatePattern("^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")]
    [string]$IPAddress
)

$RequestURL = "http://whois.arin.net/rest/ip/$IPAddress"
$Response = Invoke-RestMethod $RequestURL
$Props = [ordered]@{
    'IPAddress' = $IPAddress
    'Name' = $Response.net.Name
    'StartAddress' = $Response.net.startaddress
    'EndAddress' = $Response.net.endaddress
    'NetBlocks' = $Response.net.netBlocks.netBlock | ForEach-Object {"$($_.startaddress)/$($_.cidrLength)"}
    'RegisteredOrganization' = $Response.net.orgref.Name
    'LastUpdated' = $Response.net.updateDate -as [datetime]
}
New-Object -TypeName PSObject -Property $Props