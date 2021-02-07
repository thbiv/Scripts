If ($IsWindows) {
	$OutputPath = $Env:temp
} Else {
	$OutputPath = $Home
}

If ((Test-Path -Path $($Profile.CurrentUserAllHosts)) -eq $False) {
	New-Item -Path $($Profile.CurrentUserAllHosts) -ItemType File -Force
}

$Params = @{
	'Uri' = 'https://api.github.com/repos/thbiv/Powershell-Profile/releases/latest'
	'Headers' = @{"Accept"="application/json"}
	'Method' = 'Get'
	'UseBasicParsing' = $True
}
$Response = Invoke-RestMethod @Params
$URL = $Response.assets.browser_download_url
Invoke-WebRequest -Uri $URL -OutFile "$OutputPath\$($Response.assets.name)"
Expand-Archive -Path "$OutputPath\$($Response.assets.name)" -DestinationPath "$OutputPath" -Force -ErrorAction 'Stop'
Copy-Item -Path "$OutputPath\Profile.ps1" -Destination $($Profile.CurrentUserAllHosts) -ErrorAction 'Stop' -Force
Write-Host "Installed Powershell-Profile: $($Response.name)"