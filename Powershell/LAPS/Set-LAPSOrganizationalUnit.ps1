#Requires -Module AdmPwd.PS
[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$True)]
    [string]$Identity,

    [Parameter(Mandatory=$True)]
    [string]$ReadGroup,

    [Parameter(Mandatory=$True)]
    [string]$ResetGroup
)

If ($PSCmdlet.ShouldProcess("$Identity", "Set computer object self permission")) {
    # Adds the Write permission to ms-Mcs-AdmPwdExpirationTime and ms-Mcs-AdmPwd to SELF
    
    Try {
        Set-AdmPwdComputerSelfPermission -Identity $Identity -ErrorAction Stop
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Throw $ErrorMessage
    }
}

If ($PSCmdlet.ShouldProcess("$Identity", "Set read password permission")) {
    # Adds the CONTROL_ACCESS permission to ms-Mcs-AdmPwd attribute

    Try {
        Set-AdmPwdReadPasswordPermission -Identity $Identity -AllowedPrincipals $ReadGroup -ErrorAction Stop
        Set-AdmPwdReadPasswordPermission -Identity $Identity -AllowedPrincipals $ResetGroup -ErrorAction Stop
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Throw $ErrorMessage
    }
}

If ($PSCmdlet.ShouldProcess("$Identity", "Set reset password permission")) {
    # Adds the Write permission to ms-Mcs-AdmPwdExpirationTime attribute

    Try {
        Set-AdmPwdResetPasswordPermission -Identity $Identity -AllowedPrincipals $ResetGroup -ErrorAction Stop
    } Catch {
        $ErrorMessage = $_.Exception.Message
        Throw $ErrorMessage
    }
}