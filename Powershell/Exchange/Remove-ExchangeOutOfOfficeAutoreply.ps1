<#
.SYNOPSIS
Turns off Out-of-Office Auto-reply for a mailbox.

.PARAMETER User
Username for the account you wish to turn off OOO for.
#>

[CmdletBinding(SupportShouldProcess)]
Param (
    [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
    [Alias('SAMAccountName')]
    [string[]]$User
)
Begin {}
Process {
    ForEach ($Obj in $User) {
        $Params = @{
                    'Identity'=$Obj;
                    'AutoReplyState'='Disabled';
        }
        If ($PSCmdlet.ShouldProcess($Obj,'Disable Out-of-Office Autoreply') {
            Set-MailboxAutoReplyConfiguration @Params
        }
    }
}
End {}