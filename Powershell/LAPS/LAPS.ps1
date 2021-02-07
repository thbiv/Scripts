# Run this as Schema Admin from an FSRoot.com DC
# Creates 2 new attributes in the computer object schema
Update-AdmPwdADSchema

# Production OUs
# Adds the Write permission to ms-Mcs-AdmPwdExpirationTime and ms-Mcs-AdmPwd to SELF
Set-AdmPwdComputerSelfPermission -Identity 'OU=Workstation,OU=Container,DC=DomainName,DC=com'

# Adds the CONTROL_ACCESS permission to ms-Mcs-AdmPwd attribute
Set-AdmPwdReadPasswordPermission -Identity 'OU=Workstation,OU=Container,DC=DomainName,DC=com' -AllowedPrincipals "LAPS Workstation - Password Readers"
Set-AdmPwdReadPasswordPermission -Identity 'OU=Workstation,OU=Container,DC=DomainName,DC=com' -AllowedPrincipals "LAPS Workstation - Admins"

# Adds the Write permission to ms-Mcs-AdmPwdExpirationTime attribute
Set-AdmPwdResetPasswordPermission -Identity 'OU=Workstation,OU=Container,DC=DomainName,DC=com' -AllowedPrincipals "LAPS Workstation - Admins"