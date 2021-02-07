[cmdletBinding()]
Param (
	[string]$OutputPath = "$Home\Desktop",
	[string]$OutputFileName = ('AllUsers_Report_{0:yyyyMMdd}' -f (Get-Date)),
	[array]$OUs = @('OU=Internal Users,DC=Selene1,DC=FSRoot,DC=com','OU=External Users,DC=Selene1,DC=FSRoot,DC=com','OU=Administrators,DC=Selene1,DC=FSRoot,DC=com'),
	[switch]$ExportToExcel
)

Write-Verbose "[AllUsersReport] Script Started"
# Function to reformat the CanonicalName of the object to just the containers it is in
Function Get-ADLocation {
	Param ([string]$CanonicalName)
	$LocArr = $CanonicalName.Split('/')
	$ADLocation = ($LocArr[1..($LocArr.Count - 2)]) -join '/'
	Write-Output $ADLocation
}
# Function to convert the memberof property of an AD User account to a semicolon seperated string of group names
Function Get-GroupListArray {
	Param ($List)
	$Names = ForEach ($Line in $List) {($line -split ",*..=")[1]}
	$Output = $Names -Join ';'
	Write-Output $Output
}

If ($ExportToExcel) {
	Write-Verbose "Output will be an Excel document"
	Write-Verbose "Appending '.xlsx' to the output file name"
	$OutputFileName = "$OutputFileName"+".xlsx"
}
Else {
	Write-Verbose "Output will be a CSV file"
	Write-Verbose "Appending '.csv' to the output file name"
	$OutputFileName = "$OutputfileName"+".xsv"
}

$AllUsers = @()
ForEach ($OU in $OUs) {
	$OULabel = (($OU -split ',')[0]).TrimStart("OU=")
	Write-Verbose "[AllUsersReport][$OULabel] Processing OU: $OULabel"
	# Retrieve all objects from the provided search base
	Write-Verbose "[AllUsersReport][$OULabel] Collecting Users from OU"
	$Objects = Get-ADUser -Server SFHOUDC01 -SearchBase $OU -Filter * -Properties * | Select-Object DisplayName,SAMAccountName,whenCreated,LogonCount,LastLogonDate,CanonicalName,Description,Enabled,AccountExpirationDate,ScriptPath,MemberOf

	Write-Verbose "[AllUsersReport][$OULabel] Looking through Users and creating then adding custom object to output"
	ForEach ($Object in $Objects) {
		$ObjectLabel = $($Object.SAMAccountName)
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Processing User: $($Object.DisplayName)"
		If ($Null -eq $($Object.LastLogonDate)) {
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] 'LastLogonDate' is Null, Setting variable to 'Never'"
		$Days_LastLogon = "Never"
		} Else {
			Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Calculating number of days since last logon"
			$Days_LastLogon = 	(New-TimeSpan -Start ($Object.LastLogonDate) -End ($DateTimeStamp)).Days
		}
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Calling 'Get-ADLocation' private fuction to reformat the 'CanonicalName'"
		$ADLocation = Get-ADLocation -CanonicalName $($Object.CanonicalName)
		If ($($Object.Enabled) -eq $true) {
			Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Account Enabled"
			$AccountStatus = 'Enabled'
		}ElseIF ($($Object.Enabled) -eq $false) {
			Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Account Disabled"
			$AccountStatus = 'Disabled'
		}
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Calling 'Get-GroupListArray' private fuction to convert the 'memberof' data"
		$Groups = Get-GroupListArray -List $($Object.MemberOf)

		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Building the properties of the new custom object"
		$Props = [ordered]@{
			'DisplayName'=$($Object.DisplayName)
			'SAMAccountName'=$($Object.SAMAccountName)
			'WhenCreated'=$($Object.whenCreated)
			'LogonCount'=$($Object.LogonCount)
			'LastLogonDate'=$($Object.LastLogonDate)
			'DaysLastLogon'=$Days_LastLogon
			'ADLocation'=$ADLocation
			'AccountStatus'=$AccountStatus
			'AccountExpirationDate'=$($Object.AccountExpirationDate)
			'LogonScript'=$($Object.ScriptPath)
			'Description'=$($Object.Description)
			'Groups'=$Groups
		}
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Creating custom object and adding to the Output Array"
		$Obj = New-Object -TypeName PSObject -Property $Props
		$AllUsers += $Obj
		Write-Verbose "[AllUsersReport][$OULabel][$ObjectLabel] Processing Complete on $($Object.DisplayName)"
	}
	Write-Verbose "[AllUsersReport][$OULabel] Processing Complete on $OULabel"
}
Write-Verbose "[AllUsersReport] Sort output by 'ADLocation'-Ascending and 'DaysLastLogon'-Descending"
$SortProp1 = @{Expression='ADLocation'; Ascending=$True}
$SortProp2 = @{Expression='DaysLastLogon'; Descending=$True}
$AllUsers = $AllUsers | Sort-Object $SortProp1,$SortProp2

If ($ExportToExcel) { # Creates custom XLSX document and exports the file.
	Write-Verbose "[AllUsersReport][ExportToExcel] Creating custom XLSX document"
	$FilePath = Join-Path -Path $OutputPath -ChildPath $OutputFileName
	Write-Verbose "[AllUsersReport][ExportToExcel] Output File: $FilePath"
	Write-Verbose "[AllUsersReport][ExportToExcel] Using the ExcelPSLib Module"
	Write-Verbose "[AllUsersReport][ExportToExcel] Document Title: All_Users_Report Author: Powershell"
	[OfficeOpenXml.ExcelPackage]$ObjExcel = New-OOXMLPackage -author "Powershell" -title "All_Users_Report"
	[OfficeOpenXml.ExcelWorkbook]$xWorkBook = $ObjExcel | Get-OOXMLWorkbook
	Write-Verbose "[AllUsersReport][ExportToExcel] WorkSheet Name: Users"
	$ObjExcel | Add-OOXMLWorksheet -WorkSheetName "Users"
	$xWorkSheetUsers = $xWorkBook | Select-OOXMLWorkSheet -WorkSheetName "Users"

	Write-Verbose "[AllUsersReport][ExportToExcel] Creating Column Headers"
	$StyleHeader = New-OOXMLStyleSheet -WorkBook $xWorkBook -Name "Header" -Bold
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 1 -Value "DisplayName" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 2 -Value "SAMAccountName" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 3 -Value "WhenCreated" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 4 -Value "LogonCount" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 5 -Value "LastLogonDate" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 6 -Value "DaysLastLogon" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 7 -Value "OUName" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 8 -Value "AccountStatus" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 9 -Value "AccountExpirationDate" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 10 -Value "LogonScript" -StyleSheet $StyleHeader | Out-Null
	$xWorkSheetUsers | Set-OOXMLRangeValue -Row 1 -Col 11 -Value "Description" -StyleSheet $StyleHeader | Out-Null

	$Row = 2

	ForEach ($User in $AllUsers) {
		Write-Verbose "[AllUsersReport][ExportToExcel][$($User.SAMAccountName)] Writing user to Row $Row"
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 1 -Value $(CheckIfNullValue $($User.DisplayName)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 2 -Value $(CheckIfNullValue $($User.SAMAccountName)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 3 -Value $(CheckIfNullValue $($User.whenCreated)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 4 -Value $(CheckIfNullValue $($User.LogonCount)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 5 -Value $(CheckIfNullValue $($User.LastLogonDate)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 6 -Value $(CheckIfNullValue $($User.DaysLastLogon)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 7 -Value $(CheckIfNullValue $($User.ADLocation)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 8 -Value $(CheckIfNullValue $($User.AccountStatus)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 9 -Value $(CheckIfNullValue $($User.AccountExpirationDate)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 10 -Value $(CheckIfNullValue $($User.LogonScript)) | Out-Null
		$xWorkSheetUsers | Set-OOXMLRangeValue -Row $Row -Col 11 -Value $(CheckIfNullValue $($User.Description)) | Out-Null
		$Row++
	}
	Write-Verbose "[AllUsersReport][ExportToExcel] Autofitting all columns and Saving document to $FilePath"
	1..11 | ForEach-Object {$xWorkSheetUsers.Column($_).AutoFit()}
	$ObjExcel | Save-OOXMLPackage -FileFullPath $FilePath -Dispose
} Else {
	Write-Verbose "[AllUsersReport][ExportToCSV] Creating CSV file"
	$FilePath = Join-Path -Path $OutputPath -ChildPath $OutputFileName
	Write-Verbose "[AllUsersReport][ExportToCSV] Output File: $FilePath"
	Write-Verbose "[AllUsersReport][ExportToCSV] Saving document to $FilePath"
	$AllUsers | Export-Csv -Path $FilePath -NoTypeInformation
}
Write-Verbose "[AllUsersReport] Script Ended"