[CmdletBinding()]
Param (
    [string]$Path
)
$OUs = @()
$OutputFile = 'Not_In_SEM.csv'
$ADData = @()
ForEach ($OU in $OUs) {
    Write-Host "Adding: $OU"
    $OUData = Get-ADComputer -SearchBase $OU -Filter * -Properties * | Select-Object *
    $OUData | Add-Member -MemberType NoteProperty -Name 'Tag' -Value 0
    $ADData += $OUData
}

$SEMData = Import-Csv -Path $Path

$ADData | ForEach-Object {
    $ADComputerName = $_.Name
    Write-Host "ADComputerName: $ADComputerName"
    ForEach ($SEMItem in $SEMData) {
        $SEMComputerName = (($SEMItem."Node Name").split('.'))[0]
        If ($SEMComputerName -eq $ADComputerName) {
            $_.Tag = 1
            Write-Host "Matched: $SEMComputerName"
        }
    }
}
$ADData | Where-Object {$_.Tag -eq 0} | Select-Object Name,CanonicalName,Enabled,Description | Export-Csv -Path $OutputFile -NoTypeInformation -Force