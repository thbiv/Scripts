<#
.SYNOPSIS
Enables the Out-of-Office feature of an Exchange mailbox.

.DESCRIPTION
Enables the Out-of-Office feature of an Exchange mailbox.

.PARAMETER User
The Username for the account to set OOO on.

.PARAMETER InternalMessage
The message that will be sent to internal addresses.

.PARAMETER ExternalMessage
The message that will be sent to external addresses.

.EXAMPLE
PS C:\> $Message = @"
>> This is a test OOO Message.
>>
>> Thanks,
>> Jane Doe
>> "@
PS C:\> Set-ExchangeOutOfOfficeAutoreply.ps1 -User johnd -InternalMessage $Message -ExternalMessage $Message

This example first creates a variable that consists of a HERE string of the message to be sent.
Then using that variable in the Set-ExchangeOOO command to set the OOO for an account for John Doe.

.NOTES
Version: 1.1.4
Author: Thomas Barratt
#>
[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$True,Position=1)]
    [string[]]$User,

    [Parameter(Mandatory=$True)]
    [string]$InternalMessage,

    [Parameter(Mandatory=$True)]
    [string]$ExternalMessage
)
#Loop through the Users and set OOO
ForEach ($Obj in $User) {
    # Build parameter list to use with the cmdlet.
    $Params = @{
        'Identity'=$Obj
        'AutoReplyState'='Enabled'
        'ExternalAudience'='All'
        'InternalMessage'=$InternalMessage
        'ExternalMessage'=$ExternalMessage
    }
    # Execute the cmdlet with the parameters in $Params
    If ($PSCmdlet.ShouldProcess($Obj,'Enable Out-of-Office Autoreply')){
        Set-MailboxAutoReplyConfiguration @Params
    }
}