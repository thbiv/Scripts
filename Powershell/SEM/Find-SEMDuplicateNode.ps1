[CmdletBinding()]
Param (
    [string]$Path
)

$Nodes = Get-Content -Path $Path
$Names = @()
ForEach ($Node in $Nodes) {
    $Obj = ($Node.split('.'))[0]
    $Names += $Obj
}
$UniqueNodes = $Names | Select-Object -Unique
(Compare-Object -ReferenceObject $UniqueNodes -DifferenceObject $Names).InputObject | Select-Object -Unique