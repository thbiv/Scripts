<#
.SYNOPSIS
    Gathers which computers a specific user is logged onto.

.DESCRIPTION
    Gathers which computers a specific user is logged onto.
    Requires local administrator priveledges on each computer being checked.

.PARAMETER UserName
    Username of the AD user account to search computers for.

.PARAMETER SearchBase
    The Distinguished name of the OU to target the search.
    Default to 'OU=Workstation,OU=Systems,DC=Selene1,DC=FSRoot,DC=com'

.PARAMETER TranscriptFile
    Path and file to save the transcript of this function.
    Defaults to the user's desktop folder and a default name of Transcript.txt.

.PARAMETER OutputPath
    Path where the output of the function will be saved.
    Default to the user's desktop folder.
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$UserName,

    [Parameter(Mandatory=$False)]
    [string]$SearchBase,

    [Parameter(Mandatory=$False)]
    [string]$TranscriptFile = $(Join-Path -Path $Home -ChildPath 'Desktop\Find-LoggedOnUser_Transcript.txt'),

    [Parameter(Mandatory=$False)]
    [string]$OutputPath = $(Join-Path -Path $Home -ChildPath 'Desktop')
)

$OutputFileName = "LoggedOnComputers_{0}_{1:yyyyMMdd}.txt" -f $UserName, $(Get-Date)
$OutputFilePath = Join-Path -Path $OutputPath -ChildPath $OutputFileName
Start-Transcript -Path $TranscriptFile -Force
Write-Verbose "Searching for Username: $Username"
Write-Verbose "SearchBase: $SearchBase"
Write-Verbose "OutputFilePath: $OutputFilePath"
$Computers = Get-ADComputer -SearchBase $SearchBase -Properties * -Filter * | Where-Object {($_.Enabled -eq $True) -and ($_.OperatingSystem -like "Windows*")}
$LoggedOn = @()
ForEach ($Computer in $Computers) {
    $ComputerName = $Computer.Name
    Write-Verbose "[$ComputerName] Testing Connection"
    If ((Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) -eq $True){
        Write-Verbose "[$ComputerName] Online"
        Try {
            #$Proc = Get-WmiObject -Class Win32_Process -Computer $ComputerName -Filter "Name = 'explorer.exe'" -ErrorAction Stop
            $Proc = Get-CimInstance -ClassName Win32_Process -ComputerName $ComputerName -Filter "Name= 'explorer.exe'" -ErrorAction Stop
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Write-Warning "[$ComputerName] $ErrorMessage"
        }
        ForEach ($P in $Proc) {
            #$Temp = ($P.GetOwner()).User
            $Temp = (Invoke-CimMethod -InputObject $P -MethodName GetOwner).User
            If ($Temp -eq $UserName){
                $LoggedOn += $ComputerName
                Write-Verbose "[$ComputerName] SUCCESS: $UserName is logged on" -ForegroundColor Blue -BackgroundColor Black
            }
        }
    } Else {
        Write-Verbose "[$ComputerName] Offline" -BackgroundColor Red -ForegroundColor White
    }
    $LoggedOn | Out-File -FilePath $OutputFilePath -Force
}
Stop-Transcript