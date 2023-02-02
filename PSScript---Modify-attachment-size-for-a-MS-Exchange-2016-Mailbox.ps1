# Thank you for using my script, visit my GitHub for more https://github.com/emad-mukhtar
Write-Host "Thank you for using my script, visit my GitHub for more https://github.com/emad-mukhtar"

# This is a PowerShell script to modify inbound & outbound email attachment size
Write-Host "This is a PowerShell script to modify inbound & outbound email attachment size"

# Prompt for server name, username, password, and email address
$server = Read-Host "Enter the name of the Exchange server"
$username = Read-Host "Enter the username for connecting to the Exchange server"
$password = Read-Host -AsSecureString "Enter the password for connecting to the Exchange server"
$userName = Read-Host "Enter the email address of the single mailbox to modify"

# Import the Exchange PowerShell module and connect to the Exchange server
Import-Module ExchangeOnlineManagement
$cred = New-Object System.Management.Automation.PSCredential($username, $password)
Connect-ExchangeOnline -Credential $cred

# Prompt for the new MaxReceiveSize and MaxSendSize in MB
$newMaxReceiveSize = Read-Host "Enter the new MaxReceiveSize in MB for the single mailbox"
$newMaxSendSize = Read-Host "Enter the new MaxSendSize in MB for the single mailbox"

# Adding fail safe checks
$userQuotas = Get-Mailbox $userName | Select-Object StorageLimitStatus,ProhibitSendQuota,IssueWarningQuota,ProhibitSendReceiveQuota
if($newMaxReceiveSize -gt $userQuotas.StorageLimitStatus.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.StorageLimitStatus.Value.ToMB()){
    Write-Host "The specified max attachment size exceeds the user's mailbox size. Please choose a smaller size."
    return
}
if($newMaxReceiveSize -gt $userQuotas.ProhibitSendQuota.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.IssueWarningQuota.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.ProhibitSendReceiveQuota.Value.ToMB()){
    Write-Host "The specified max attachment size exceeds the user's ProhibitSend, IssueWarning, or ProhibitSendReceive quotas. Please choose a smaller size."
    return
}

# Prompt for confirmation
Write-Host "You are about to modify the max attachment size for $userName. MaxReceiveSize to $newMaxReceiveSize MB, and MaxSendSize to $newMaxSendSize MB"
$confirm = Read-Host "Please type OK to confirm, or type CANCEL to cancel the operation"
if ($confirm -ne "OK"){
    Write-Host "Operation canceled by user"
    return
}

# Modify the MaxReceiveSize and MaxSendSize for the specified user
Set-Mailbox $userName -MaxReceiveSize $newMaxReceiveSize'MB' -MaxSendSize $newMaxSendSize'MB'
Write-Host "MaxReceiveSize and MaxSendSize for $userName have been successfully modified to $newMaxReceiveSize MB and $newMaxSendSize MB, respectively"
Write-Host "You can always run the following command to check the currently configured Attachment Size: Get-Mailbox -Identity "<UserName>" | Select MaxReceiveSize, MaxSendSize"
