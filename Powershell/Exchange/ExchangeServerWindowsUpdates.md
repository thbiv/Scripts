# Perform Windows Updates on Exchange Server Environment

1. Verify that all database replicas are healthy and that all 'Mounted' databases are on Houston servers

    ```Powershell
    Clear-Host;Get-MailboxDatabaseCopyStatus * | Sort-Object Name
    ```

1. Disable the 'Constant DAG Monitoring' task from Utility Workstation #3

    ```Powershell
    $Session = New-CimSession -ComputerName UTIL03
    Disable-ScheduledTask -CimSession $Session -TaskName 'Constant DAG Monitoring'
    ```

1. Install Updates on all DR Exchange servers at the same time.

1. Reboot the DR exchange servers at 5 minute intervals. Wait for all servers to boot up fully

1. Wait for replay queues to clear

1. Move all the active mailbox databases from the Primary Site server to their DR server equivalent and allow the replay queues to clear before moving on.

    ```Powershell
    Move-ActiveMailboxDatabase -Server EX01 -ActivateOnServer EXDR01
    Move-ActiveMailboxDatabase -Server EX02 -ActivateOnServer EXDR02
    Move-ActiveMailboxDatabase -Server EX03 -ActivateOnServer EXDR03
    ```

1. Install Updates on all Primary Site servers at the same time

1. Reboot the Primary Site servers at 5 minute intervals and wait for the replay queues to clear

1. Move all the active mailbox databases back to the Primary Site servers then wait for the replay queues to clear

    ```Powershell
    Move-ActiveMailboxDatabase -Server EXDR01 -ActivateOnServer EX01
    Move-ActiveMailboxDatabase -Server EXDR02 -ActivateOnServer EX02
    Move-ActiveMailboxDatabase -Server EXDR03 -ActivateOnServer EX03
    ```

1. Enable the Constant DAG Monitor Task from Utility Workstation #3

    ```Powershell
    Enable-ScheduledTask -CimSession $Session -TaskName 'Constant DAG Monitoring'
    Remove-CimSession -CimSession $Session
    ```