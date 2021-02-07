# $Server - The name of the domain controller that the changes will be executed on.
$Server = ''

# $UserListCSV - Path to the CSV input file containing information for the new user accounts.
#		The column headers must be set exact.
#			UserName,Password,FirstName,LastName,Name,Office,EmailAddress,PhoneNumber,StreetAddress,City,State,ZipCode,Company,Manager
$UserListCSV = "$Home\Desktop\NewAccountList.csv"

# $OU - distinguished name of the organizational unit where the accounts will be created.
$OU = ''

# $Initials - Initials of the person who is running this script.
$Initials = ''

# $RequestNumber - ticket number that pertains to the creation of these accounts
$RequestNumber = ''

#####################################################################################
############### DO NOT CHANGE ANYTHING BELOW THIS POINT #############################
#####################################################################################

$NewDescription = "Req {0:MM/dd/yyy} {1} {2}" -f $(Get-Date), $RequestNumber, $Initials
$Users = Import-Csv -Path $UserListCSV
ForEach ($User in $Users) {
	Try{
		$Password = (ConvertTo-SecureString $($User.Password) -Force)
		$Props = @{
			'Path'=$OU
			'Server'=$Server
			'Enabled'=$True
			'SAMAccountName'=$($User.UserName)
			'UserPrincipalName'="$($User.UserName)@selene1.fsroot.com"
			'Name'=$($User.Name)
			'DisplayName'=$($User.Name)
			'GivenName'=$($User.FirstName)
			'Surname'=$($User.LastName)
			'Company'=$($User.Company)
			'Manager'=$($User.Manager)
			'Office'=$($User.Office)
			'EmailAddress'=$($User.EmailAddress)
			'OfficePhone'=$($User.PhoneNumber)
			'StreetAddress'=$($User.StreetAddress)
			'City'=$($User.City)
			'State'=$($User.State)
			'PostalCode'=$($User.ZipCode)
			'AccountPassword'=$Password
			'Description'=$NewDescription
		}
		New-ADUser @Props -ErrorAction STOP
		Write-Information -MessageData "[$($User.UserName)] Account created for '$($User.Name)' successfully" -InformationAction Continue
	}
	Catch {
		$ErrorMessage = $_.exception.message
		Write-Warning "[$($User.UserName)] $ErrorMessage"
	}
}