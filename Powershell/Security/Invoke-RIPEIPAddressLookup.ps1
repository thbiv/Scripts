<#
Work In Progress
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True,Position=0)]
    [ValidatePattern("^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")]
    [string]$IPAddress
)

$RequestURL = "http://rest.db.ripe.net/search?query-string=$IPAddress"
$Response = Invoke-RestMethod $RequestURL
$Obj_Inetnum = $Response.'whois-resources'.objects.object | where-object { $_.type -eq 'inetnum'}
$Obj_Org = $Response.'whois-resources'.objects.object | where-object { $_.type -eq 'organisation'}
$Obj_Role = $Response.'whois-resources'.objects.object | where-object { $_.type -eq 'role'}

$IPBlock = ((($Obj_Inetnum.attributes.attribute | Where-Object {$_.name -eq 'inetnum'}).value).split('-')).trim()

$Props = [ordered]@{
    'IPAddress' = $IPAddress
    'Name' = ($Obj_Org.attributes.attribute | Where-Object {$_.name -eq 'org-name'}).value
    'StartAddress' = $($IPBlock[0])
    'EndAddress' = $($IPBlock[1])
}
