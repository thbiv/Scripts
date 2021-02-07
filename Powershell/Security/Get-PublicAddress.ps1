$Result = Invoke-RestMethod -Uri http://ipinfo.io/json
$Props = [ordered]@{
    IPAddress = $Result.ip
    Hostname = $Result.hostname
    City = $Result.city
    Region = $Result.region
    Country = $Result.country
    Location = $Result.loc
    Organization = $Result.org
    PostalCode = $Result.postal
    TimeZone = $Result.timezone
    ReadMe = $Result.readme
}
Write-Output $(New-Object -TypeName PSObject -Property $Props)