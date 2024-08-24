# Exchange Mailbox Attachment Size Modifier
# Author: Emad Mukhtar (https://github.com/emad-mukhtar)

function Test-ExchangeConnection {
    try {
        Get-ExchangeServer -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Get-SecureInput {
    param ([string]$prompt)
    $secureString = Read-Host -Prompt $prompt -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

function Test-PositiveInteger {
    param ([string]$value)
    return $value -match '^\d+$' -and [int]$value -gt 0
}

Write-Host "Exchange Mailbox Attachment Size Modifier"
Write-Host "For more scripts, visit: https://github.com/emad-mukhtar"

# Prompt for connection details
$server = Read-Host "Enter the name of the Exchange server"
$username = Read-Host "Enter the username for connecting to the Exchange server"
$password = Get-SecureInput "Enter the password for connecting to the Exchange server"

# Import module and connect to Exchange
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $cred = New-Object System.Management.Automation.PSCredential($username, (ConvertTo-SecureString $password -AsPlainText -Force))
    Connect-ExchangeOnline -Credential $cred -ErrorAction Stop
}
catch {
    Write-Host "Failed to connect to Exchange: $_" -ForegroundColor Red
    exit
}

if (-not (Test-ExchangeConnection)) {
    Write-Host "Failed to establish a connection to Exchange." -ForegroundColor Red
    exit
}

# Prompt for mailbox and size details
$userName = Read-Host "Enter the email address of the mailbox to modify"

do {
    $newMaxReceiveSize = Read-Host "Enter the new MaxReceiveSize in MB"
} while (-not (Test-PositiveInteger $newMaxReceiveSize))

do {
    $newMaxSendSize = Read-Host "Enter the new MaxSendSize in MB"
} while (-not (Test-PositiveInteger $newMaxSendSize))

# Fetch mailbox quotas
try {
    $userQuotas = Get-Mailbox $userName -ErrorAction Stop | 
                  Select-Object StorageLimitStatus, ProhibitSendQuota, IssueWarningQuota, ProhibitSendReceiveQuota
}
catch {
    Write-Host "Failed to fetch mailbox details: $_" -ForegroundColor Red
    exit
}

# Validate new sizes against quotas
$quotaErrors = @()
if ([int]$newMaxReceiveSize -gt $userQuotas.StorageLimitStatus.Value.ToMB()) {
    $quotaErrors += "MaxReceiveSize exceeds the mailbox size limit."
}
if ([int]$newMaxReceiveSize -gt $userQuotas.ProhibitSendReceiveQuota.Value.ToMB()) {
    $quotaErrors += "MaxReceiveSize exceeds the ProhibitSendReceiveQuota."
}
if ([int]$newMaxSendSize -gt $userQuotas.ProhibitSendQuota.Value.ToMB()) {
    $quotaErrors += "MaxSendSize exceeds the ProhibitSendQuota."
}

if ($quotaErrors.Count -gt 0) {
    Write-Host "The following errors were encountered:" -ForegroundColor Red
    $quotaErrors | ForEach-Object { Write-Host "- $_" -ForegroundColor Red }
    exit
}

# Prompt for confirmation
Write-Host "You are about to modify the max attachment size for $userName:"
Write-Host "MaxReceiveSize: $newMaxReceiveSize MB"
Write-Host "MaxSendSize: $newMaxSendSize MB"
$confirm = Read-Host "Type 'YES' to confirm, or anything else to cancel"

if ($confirm -ne "YES") {
    Write-Host "Operation canceled by user." -ForegroundColor Yellow
    exit
}

# Modify mailbox settings
try {
    Set-Mailbox $userName -MaxReceiveSize "$newMaxReceiveSize MB" -MaxSendSize "$newMaxSendSize MB" -ErrorAction Stop
    Write-Host "Mailbox settings updated successfully." -ForegroundColor Green
    Write-Host "To verify the changes, run:"
    Write-Host "Get-Mailbox -Identity `"$userName`" | Select-Object MaxReceiveSize, MaxSendSize" -ForegroundColor Cyan
}
catch {
    Write-Host "Failed to update mailbox settings: $_" -ForegroundColor Red
}

# Disconnect from Exchange
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
