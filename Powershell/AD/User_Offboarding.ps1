[CmdletBinding()]
Param()

$StopWatch = [system.diagnostics.stopwatch]::startNew()
###########################################################################################################################
# UserListEnabled - a Variable to control where the script retrieves its input from.
#	$True - The script uses a text file located in the UserListPath variable to input a list of usernames,
#		one username per line.
#	$False - The script uses the Users variable to input one or a few usernames if you do not want to use the text file.
#		Please refer to the comment for the Users variable for more information.
$UserListEnabled = $False

# Input Variables
# $Users - comma separated array of usernames.
#	Example 1 - Single name
#		@('janed')
#	Example 2 - Multiple names
#		@('janed','johnd')
#	Not used if $UserListEnabled is set to $True
$Users = @('')

# $UserListPath - The path to the text file where the script will get a list of usernames, one per line.
#	Not used if $UserListEnabled is set to $False
$UserListPath = "$Home\Desktop\DisableList.txt"
###########################################################################################################################

# SetDescription Section - Appends description to the user account.
#	$True - the script will append a description to the user account.
#	$False - the section will be skipped
#	This section should always be run (set to $True).
$SetDescription = $True

# Varialbes used in the SetDescription section.
# $Initials - The initials of the technician running the script.
$Initials = ''

# $RequestNumber - the request number of for the ticket that was created for disabling the account.
$RequestNumber = ''

# AppendDescription - This variable uses the $Initials and $RequestNumber variables to build the appended descirption text.
#	No need to change this
$AppendDescription = "Disabled {0:MM/dd/yyy} {1} {2}" -f $(Get-Date), $RequestNumber, $Initials
###########################################################################################################################

# ExportSecurityGroups Section - Exports groups for later reference if needed.
#	$True - The script will capture all security/distribution groups in the MemberOf property of the account
#		along with other specific information and create an excel file containing the information.
#	$False - The script will skip this section.
#	Please Note: DO NOT USE THIS SECTION FOR CITRIX USERS. NO NEED TO REFERENCE THEIR GROUPS AS THE TECH FORM INCLUDES EVERYTHING
$ExportSecurityGroups = $True

# Variables used in the ExportSecurityGroups section.
# $ExportGroupsFilePath - Path to where the excel file will be copied to.
#	Do not change unless instructed by the Security Team.
$ExportGroupsFilePath = '\\sfhousanp01\IT\Security Team\AD\Terms'
###########################################################################################################################

# RemoveSecurityGroups Section - Removes all security groups from the AD User account except for the primary group (usually Domain Users)
#	$True - The script will remove the security groups.
#	$False - The script will not remove the groups.
#	NOTE: THIS SHOULD ALWAYS BE SET TO $True
$RemoveSecurityGroups = $True

# This section does not have associated variables
###########################################################################################################################

# $ClearProperties Section - This section clears specific properties of the AD User Account that are important to clear.
#	$True - The script will clear the properties.
#	$False - The script will skip this section.
#	Properties Cleared: TelephoneNumber,mobile,IPPhone,HomePhone,facsimileTelephoneNumber,Manager
$ClearProperties = $True

# This seciton does not have associated variables
###########################################################################################################################

# HideAddress Section - This section will hide the email address from the Global Address List.
#	$True - The script will hide the address.
#	$False - The Script will skip this section.
#	This section requires powershell to be connected to the exchange environment using a PSSession before running the script.
#	If this is set to $False, then there is not need to connect to exchange.
#	Please Note: DO NOT USE THIS SECTION FOR CITRIX USERS. THEY DO NOT HAVE a @SeleneFinance.com EMIAL ADDRESS TO HIDE
$HideAddress = $False

# This seciton does not have associated variables
###########################################################################################################################

# $DisableAndMove Section - This section will disable the account and move it to the designated OU.
#	$True - Disables the account and moves the account to the OU specified in the $DisableOU variable.
#	$False - The script will skip this section
#	THIS SHOULD ALWAYS BE SET TO $True.
$DisableAndMove = $True

# Variables used in the $DisableAndMove section.
# $DisableOU - The path to the OU where Disabled accounts are moved to.
#	Please give the Distinguished Name of the OU.
#	DO NOT CHANGE UNLESS INSTRUCTED BY THE SECURITY TEAM
$DisableOU = 'OU=To Be Deleted,DC=Selene1,DC=FSRoot,DC=com'
###########################################################################################################################

# General Variables used through out the script
# $TranscriptPath - Path to where you would like the transcript of this script execution placed.
#	The output of the script is sent to this file along with the console host itslef in case the output is more than the console buffer can handle.
#	Is a .TXT file.
#	The file will get overwritten everytime the script is executed.
$TranscriptPath = "$Home\Desktop\DisableADUserAccount_Transcript.txt"

# $Server - The Name of the domain controller where all of the changes will take place.
$Server = 'SFHOUDC01'
###########################################################################################################################
###########################################################################################################################
# Please do not change anything below this point. Doing so could make the script not run correctly.

# Function to check is a variable is Null. If yes, set it to an empty string ("")
Function CheckIfNullValue {
	Param ($Data)
	If ($Null -eq $Data) {$Output = ""}
	Else {$Output = $Data}
	Write-Output $Output
}
Start-Transcript -Path $TranscriptPath -Force
If ($UserListEnabled -eq $True) {
	$Users = Get-Content -Path $UserListPath
}
ForEach ($User in $Users) {
	$UserObject = Get-ADUser $User -Server $Server -Properties *
	If ($SetDescription -eq $True){
		Try{
			$UserObject | ForEach-Object {Set-ADUser $_ -Server $Server -ErrorAction 'Stop' -Description "$AppendDescription $($_.Description)"}
			Write-Information -MessageData "[$User] Description Appended with: $AppendDescription" -InformationAction Continue
		} Catch {
			$ErrorMessage = $_.exception.message
			Write-Warning "[$User][SetDescription] $ErrorMessage"
		}
	}
	If ($ExportSecurityGroups -eq $True){
		$GroupDNList = $UserObject | Select-Object -ExpandProperty MemberOf
		$GroupDNList = $GroupDNList | Sort-Object
		$Manager = ($($UserObject.Manager) -split ",*..=")[1]
		$FileName = "{0}_{1:yyyyMMdd}.xlsx" -f $($UserObject.DisplayName),$(Get-Date)
		[OfficeOpenXml.ExcelPackage]$ObjExcel = New-OOXMLPackage -author "Powershell" -title "SecurityGroups"
		[OfficeOpenXml.ExcelWorkbook]$xWorkBook = $ObjExcel | Get-OOXMLWorkbook
		$ObjExcel | Add-OOXMLWorksheet -WorkSheetName "Groups"
		$xWorkSheetGroups = $xWorkBook | Select-OOXMLWorkSheet -WorkSheetName "Groups"
		$StyleHeader = New-OOXMLStyleSheet -WorkBook $xWorkBook -Name "Header" -Bold
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 1 -Col 1 -Value "Name" -StyleSheet $StyleHeader | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 2 -Col 1 -Value $($UserObject.DisplayName) | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 2 -Col 3 -Value "Employee ID" -StyleSheet $StyleHeader | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 2 -Col 4 -Value $($UserObject.EmployeeID) | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 3 -Col 3 -Value "Description" -StyleSheet $StyleHeader | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 3 -Col 4 -Value $($UserObject.Description) | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 4 -Col 3 -Value "Manager" -StyleSheet $StyleHeader | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 4 -Col 4 -Value $($Manager) | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 5 -Col 3 -Value "Department" -StyleSheet $StyleHeader | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 5 -Col 4 -Value $($UserObject.Department) | Out-Null
		$xWorkSheetGroups | Set-OOXMLRangeValue -Row 4 -Col 1 -Value "Member Of" -StyleSheet $StyleHeader | Out-Null
		$Row = 5
		ForEach ($line in $GroupDNList) {
			$Group = ($line -split ",*..=")[1]
			$Type = (Get-ADGroup -Identity $line -Properties * | Select-Object groupType).groupType
			Switch ($Type) {
				"8" {$TypeName = "Dist List-Universal"}
				"-2147483640" {$TypeName = "Security Group-Universal"}
				"2" {$TypeName = "Dist List-Global"}
				"-2147483646" {$TypeName = "Security Group-Global"}
				"4" {$TypeName = "Dist List-Domain Local"}
				"-2147483644" {$TypeName = "Security Group-Domain Local"}
				default {$TypeName = "Not Available"}
			}
			$xWorkSheetGroups | Set-OOXMLRangeValue -Row $Row -Col 1 -Value $Group | Out-Null
			$xWorkSheetGroups | Set-OOXMLRangeValue -Row $Row -Col 2 -Value $TypeName | Out-Null
			$Row++
		}
		1..3 | ForEach-Object {$xWorkSheetGroups.Column($_).AutoFit()}
		$ObjExcel | Save-OOXMLPackage -FileFullPath $(Join-Path -Path $ExportGroupsFilePath -ChildPath $FileName) -Dispose
		Write-Information -MessageData "[$User] Exported All Security Groups to: $(Join-Path -Path $ExportGroupsFilePath -ChildPath $FileName)" -InformationAction Continue
	}
	If ($RemoveSecurityGroups -eq $True) {
		Try {
			Get-ADPrincipalGroupMembership -Identity $User -Server $Server | Where-Object {$_.Name -ne "Domain Users"} | ForEach-Object {
				Write-Information -MessageData "[$User] Removing: $_" -InformationAction Continue
				Remove-ADPrincipalGroupMembership -Identity $User -MemberOf $_ -Server $Server -Confirm:$false
			} | Out-Null
			Write-Information -MessageData "[$User] Removed All Security Groups" -InformationAction Continue
		} Catch {
			$ErrorMessage = $_.exception.Message
			Write-Warning "[$User][RemoveSecurityGroups] $ErrorMessage"
		}
	}
	If ($ClearProperties -eq $True){
		Try {
			$UserObject | Set-ADUser -Clear TelephoneNumber,mobile,IPPhone,HomePhone,facsimileTelephoneNumber,Manager -Server $Server -ErrorAction 'Stop'
		    Write-Information -MessageData "[$User] Cleared Properties: TelephoneNumber,mobile,IPPhone,HomePhone,facsimileTelephoneNumber,Manager" -InformationAction Continue
		} Catch {
			$ErrorMessage = $_.exception.message
			Write-Warning "[$User][ClearProperties] $ErrorMessage"
		}
	}
	If ($HideAddress -eq $True){
		Try {
			Set-Mailbox -Identity $User -HiddenFromAddressListsEnabled $True -ErrorAction 'Stop'
			Write-Information -MessageData "[$User] Mailbox Hidden from GAL" -InformationAction Continue
		} Catch {
			$ErrorMessage = $_.exception.message
			Write-Warning "[$User][HideAddress] $ErrorMessage"
		}
	}
	If ($DisableAndMove -eq $True){
		Try {
			$UserObject | Disable-ADAccount -Server $Server -ErrorAction 'Stop'
			$UserObject | Move-ADObject -TargetPath $DisableOU -Server $Server -ErrorAction 'Stop'
			Write-Information -MessageData "[$User] Disabled and Moved" -InformationAction Continue
		} Catch {
			$ErrorMessage = $_.exception.message
			Write-Warning "[$User][DisableAndMove] $ErrorMessage"
		}
	}

}
$StopWatch.Stop()
Write-Information -MessageData "[Disable-ADUserAccount] Script Execution Time: $($StopWatch.Elapsed.Days) Days $($StopWatch.Elapsed.Hours) Hours $($StopWatch.Elapsed.Minutes) Minutes $($StopWatch.Elapsed.Seconds) Seconds $($StopWatch.Elapsed.Milliseconds) Milliseconds" -InformationAction Continue
Stop-Transcript