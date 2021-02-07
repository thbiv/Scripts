@{
    SourceDLLPath = ''
    SourceBlacklistPath = ''
    IncludeAllDomainControllers = $False
    IncludeComputerNames = @('localhost')
    SendMail = $False
    MailSettings = @{
        SMTPServer = ''
        Port = ''
        To = ''
        From = ''
    }
}