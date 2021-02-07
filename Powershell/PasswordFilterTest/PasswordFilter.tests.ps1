Param (
    [Parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceDLLPath,

    [Parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
    [array]$Servers
)

Describe "Remote Version" {    
    Context "Latest Available Version" {
        $LatestAvailableVersion = $($(Get-PassFiltExLatestRelease).Version)
        $ProductionVersion = $((Get-ItemProperty -Path $SourceDLLPath | Select-Object -ExpandProperty VersionInfo).ProductVersion)
        It "Is the latest version available" {
            $ProductionVersion | Should -Be $LatestAvailableVersion
        }
    }
}
Describe 'Password Filter State by Domain Controller' {
    ForEach ($Server in $Servers) {
        Context "$Server" {
            $Result = Get-PassFiltExStatus -ServerName $Server
            It "Blacklist file should exist" {
                $Result.BlackListExists | Should -Be $True
            }
            It "Blacklist file matches the master file" {
                $Result.BlackListCurrent | Should -Be $True
            }
            It "DLL file exists" {
                $Result.DLLExists | Should -Be $True
            }
            It "Dll file upgrade needed" {
                $Result.UpgradeNeeded | Should -Be $False
            }
            It "Password filter should be enabled in the resigry" {
                $Result.Enabled | Should -Be $True
            }
        }
    }
}