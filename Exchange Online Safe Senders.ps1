# Set the log file path
$logFilePath = "$PSScriptRoot\log.txt"

# Function to write log messages to the log file
function Write-Log {
    param (
        [string]$Message
    )

    $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    $logMessage | Out-File -FilePath $logFilePath -Append
}

# Check if the ExchangeOnlineManagement module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    # Logging: Module not installed
    Write-Log -Message "The ExchangeOnlineManagement module is not installed."

    # Prompt the user to install the module
    $installModule = Read-Host "Do you want to install the ExchangeOnlineManagement module? (Y/N)"
    if ($installModule -ne 'Y' -and $installModule -ne 'y') {
        # Logging: Module install cancelled.
        Write-Log -Message "Module installation cancelled. Exiting the script."
        return
    }

    # Install the ExchangeOnlineManagement module
    try {
        Write-Log -Message "Installing the ExchangeOnlineManagement module..."
        Install-Module -Name ExchangeOnlineManagement -Force
        Write-Log -Message "The ExchangeOnlineManagement module has been installed."
        Import-Module -Name ExchangeOnlineManagement
    }
    catch {
        Write-Log -Message "Failed to install the ExchangeOnlineManagement module. Please install it manually and re-run the script."
        Write-Host "Failed to install the ExchangeOnlineManagement module. Please install it manually and re-run the script."
        return
    }
}
# Prompt the user to enter the Microsoft Tenant Name
$maxRetryCount = 3
$retryCount = 0
do {
    $strOrganization = Read-Host "Enter the Microsoft tenant name (without the '.onmicrosoft.com' suffix)"
    # Append the domain suffix to the Microsoft Tenant Name
    $strTenant = "$($strOrganization).onmicrosoft.com"
    # Input validation: Check if the input is empty or contains invalid characters
    if ([string]::IsNullOrEmpty($strOrganization) -or $strOrganization -notmatch "^[a-zA-Z0-9-]+$") {
        Write-Host "Invalid input. Microsoft Tenant Name cannot be empty or contain special characters."
    }
    else {
        # Check if the tenant name is valid
        try {
            $ipAddresses = Resolve-DnsName -Name $strTenant -ErrorAction Stop
            if ($ipAddresses) {
                break
            }
            else {
                Write-Log -Message "Invalid Microsoft tenant name. Please check the tenant name and try again."
                Write-Host "Invalid Microsoft tenant name. Please check the tenant name and try again."
                $strOrganization = $null
            }
        }
        catch {
            Write-Log -Message "An error occurred while attempting to connect to '$strTenant'. Please try again."
            Write-Host "An error occurred while attempting to connect to '$strTenant'. Please try again."
            $strOrganization = $null
        }
    }
    $retryCount++
} while ($retryCount -lt $maxRetryCount)

if ($retryCount -ge $maxRetryCount) {
    Write-Log -Message "Failed to provide a valid Microsoft tenant name after $maxRetryCount attempts. Exiting the script."
    Write-Host "Failed to provide a valid Microsoft tenant name after $maxRetryCount attempts. Exiting the script."
}

# Display information
Write-Host "Authenticating to $strTenant..."

# Authenticate to Exchange Online with a timeout of 30 seconds
try {
    $connectionTimeout = New-TimeSpan -Seconds 30
    $connectionTimer = [Diagnostics.Stopwatch]::StartNew()
    do {
        try {
            Connect-ExchangeOnline -DelegatedOrganization $strTenant
            break
        }
        catch {
            # Handle authentication errors
            if ($connectionTimer.Elapsed -gt $connectionTimeout) {
                Write-Log -Message "Failed to authenticate to Exchange Online. Please check your credentials and the tenant name."
                Write-Host "Failed to authenticate to Exchange Online. Please check your credentials and the tenant name."
                return
            }
        }
        Start-Sleep -Milliseconds 500
    } while ($true)
    
}
finally {
    $connectionTimer.Stop()
}

# Set the transport rule name
$ruleName = "Firewell Technology Solutions Safe Senders"

# Read the CSV file with the domain safe list
$csvFilePath = Join-Path $PSScriptRoot "Safe Domains.csv"
if (-not (Test-Path $csvFilePath)) {
    Write-Log -Message "CSV file 'Safe Domains.csv' not found in the script directory."
    Write-Host "CSV file 'Safe Domains.csv' not found in the script directory."
    return
}

# Retrieve the existing transport rule, if it exists
$existingRule = Get-TransportRule -Identity $ruleName -ErrorAction SilentlyContinue

if ($existingRule) {
    # Logging: Existing rule found
    Write-Log -Message "The transport rule '$ruleName' already exists. Appending domain names if not already present."

    # Retrieve the existing domain names from the rule
    $existingDomains = $existingRule.SenderDomainIs

    $newDomains = Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty Domain
    
    # Filter out the domain names that are already present in the rule
    $uniqueDomains = $newDomains | Where-Object { $existingDomains -notcontains $_ }

    if ($uniqueDomains) {
        # Logging: Domain names to be added
        Write-Log -Message "The following domain names will be added to the transport rule at '$strTenant':"
        Write-Log -Message $($uniqueDomains -join ', ')
        Write-Host "The following domain names will be added to the transport rule at '$strTenant':"
        Write-Host $($uniqueDomains -join ', ')

        # Add the new domain names to the existing rule
        $existingRule.Conditions.SenderDomainIs += $uniqueDomains

        # Update the rule
        Set-TransportRule -Identity $existingRule.Identity -Conditions $existingRule.Conditions

        # Logging: Rule updated
        Write-Log -Message "The transport rule has been updated with the new domain names."
        Write-Host "The transport rule has been updated with the new domain names."
    }
    else {
        # Logging: No domain names to add
        Write-Log -Message "No new domain names to add to the transport rule at '$strTenant'."
        Write-Host "No new domain names to add to the transport rule at '$strTenant'."
    }
}

elseif ($null = $existingrule) {
# Read safe domains from CSV file
$safeDomains = Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty Domain

if (-not $safeDomains) {
    Write-Log -Message "No safe domains found in the CSV file."
    Write-Host "No safe domains found in the CSV file."
    return
}

# Create the transport rule for safe senders
Write-Output "Creating transport rule '$ruleName'..."
New-TransportRule -Name $ruleName `
    -Comments "Authenticated emails from these domains should never be filtered or blocked" `
    -Enabled $true `
    -HasNoClassification $false `
    -HeaderContainsMessageHeader Authentication-Results `
    -HeaderContainsWords "dmarc=pass", "dmarc=bestguesspass" `
    -Mode Enforce `
    -RecipientAddressType Resolved `
    -RuleErrorAction Ignore `
    -RuleSubType None `
    -SenderAddressLocation Header `
    -SenderDomainIs $safeDomains `
    -SetHeaderName X-ETR `
    -SetHeaderValue "Bypass spam filtering for authenticated sender" `
    -SetSCL -1

Write-Log -Message "The following domains were added to the Firewell Technology Solutions 'Safe Senders' list at '$strTenant' under the rule named '$ruleName':"
Write-Log -Message $($safeDomains -join ', ')
Write-Output "The following domains were added to the Firewell Technology Solutions 'Safe Senders' list at '$strTenant':"
Write-Output $($safeDomains -join ', ')
}