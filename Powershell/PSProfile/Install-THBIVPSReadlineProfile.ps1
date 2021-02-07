If ($IsWindows) {
	$OutputPath = $Env:temp
} Else {
	$OutputPath = $Home
}

If (Test-Path -Path "$(Split-Path -Path $($Profile.CurrentUserAllHosts) -Parent)\ProfileInclude") {
    New-Item -Path "$(Split-Path -Path $($Profile.CurrentUserAllHosts) -Parent)\ProfileInclude" -ItemType Directory -Force
}

$Params = @{
	'Uri' = 'https://api.github.com/repos/thbiv/PSReadline-Profile/releases/latest'
	'Headers' = @{"Accept"="application/json"}
	'Method' = 'Get'
	'UseBasicParsing' = $True
}
$Response = Invoke-RestMethod @Params
$URL = $Response.assets.browser_download_url
Invoke-WebRequest -Uri $URL -OutFile "$OutputPath\$($Response.assets.name)"
Expand-Archive -Path "$OutputPath\$($Response.assets.name)" -DestinationPath "$OutputPath" -Force -ErrorAction 'Stop'
Copy-Item -Path "$OutputPath\PSReadline-Profile.ps1" -Destination "$(Split-Path -Path $($Profile.CurrentUserAllHosts) -Parent)\ProfileInclude\PSReadline-Profile.ps1" -ErrorAction 'Stop' -Force
Write-Host "Installed PSReadline-Profile: $($Response.name)"