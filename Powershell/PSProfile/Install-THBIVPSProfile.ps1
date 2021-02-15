If ($IsWindows) {
	$OutputPath = $Env:temp
} Else {
	$OutputPath = $Home
}

If ((Test-Path -Path $($Profile.CurrentUserAllHosts)) -eq $False) {
	New-Item -Path $($Profile.CurrentUserAllHosts) -ItemType File -Force
}
$ProfileIncludePath = "$(Split-Path -Path $($Profile.CurrentUserAllHosts) -Parent)\ProfileInclude"
If (Test-Path -Path $ProfileIncludePath) {
    New-Item -Path $ProfileIncludePath -ItemType Directory -Force
}

$PorfileParams = @{
	'Uri' = 'https://api.github.com/repos/thbiv/Powershell-Profile/releases/latest'
	'Headers' = @{"Accept"="application/json"}
	'Method' = 'Get'
	'UseBasicParsing' = $True
}
$Response = Invoke-RestMethod @PorfileParams
$URL = $Response.assets.browser_download_url
Invoke-WebRequest -Uri $URL -OutFile "$OutputPath\$($Response.assets.name)"
Expand-Archive -Path "$OutputPath\$($Response.assets.name)" -DestinationPath "$OutputPath" -Force -ErrorAction 'Stop'
Copy-Item -Path "$OutputPath\Profile.ps1" -Destination $($Profile.CurrentUserAllHosts) -ErrorAction 'Stop' -Force
Write-Host "Installed Powershell-Profile: $($Response.name)"

$PSRParams = @{
	'Uri' = 'https://api.github.com/repos/thbiv/PSReadline-Profile/releases/latest'
	'Headers' = @{"Accept"="application/json"}
	'Method' = 'Get'
	'UseBasicParsing' = $True
}
$Response = Invoke-RestMethod @PSRParams
$URL = $Response.assets.browser_download_url
Invoke-WebRequest -Uri $URL -OutFile "$OutputPath\$($Response.assets.name)"
Expand-Archive -Path "$OutputPath\$($Response.assets.name)" -DestinationPath "$OutputPath" -Force -ErrorAction 'Stop'
Copy-Item -Path "$OutputPath\PSReadline-Profile.ps1" -Destination "$(Split-Path -Path $($Profile.CurrentUserAllHosts) -Parent)\ProfileInclude\PSReadline-Profile.ps1" -ErrorAction 'Stop' -Force
Write-Host "Installed PSReadline-Profile: $($Response.name)"