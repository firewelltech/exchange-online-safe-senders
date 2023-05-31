# Set the log file path
$logFilePath = "$PSScriptRoot\log.txt"

# Set the transport rule name. Change to suit your needs. The script will check to see if this exact rule exists,
# and if this rule does not exist, it will create the rule; if the rule does exist, it will overwrite it.
$ruleName = "Firewell Technology Solutions Safe Senders"

# Define custom functions
# ================================================================================
# Function to write log messages to the log file
function Write-Log {
    param (
        [string]$Message
    )

    $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    $logMessage | Out-File -FilePath $logFilePath -Append -Encoding utf8
}

# ================================================================================
# Function to check for required modules and install them if they're not already installed
function Install-RequiredModules {
    param (
        [array]$Modules
    )
    foreach ($module in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Log -Message "The $($module) module is not installed."
            try {
            Install-Module -Name $module -force -scope CurrentUser -ErrorAction Stop
            }
            catch {
                Write-Log -Message "$($module) failed to install. Please install it manually and re-run the script."
                Write-Host "$($module) failed to install. Please install it manually and re-run the script."
                return
            }
            Write-Log -Message "$($module) module installed successfully."
        }
        else {
            Write-Log -Message "$($module) already installed."
        }
    }
}

# ================================================================================
# Function returns a list of tenant domains from Partner Center that have Microsoft 365 or Exchange Online subscriptions
function Get-PartnerTenants {
    $exchangeOnlineTenants = @()
    # Authenticate with Partner Center credentials
    try {
        Connect-PartnerCenter -ErrorAction Stop
    }
    catch {
        $errorCode = "$($_.Exception.HResult): $($_.Exception.ErrorCode)"
        Write-Log -Message "Failed to connect to Partner Center: $($errorCode)"
        Write-Host "Failed to connect to Partner Center: $($errorCode)"
        return
    }   
    # Retrieve the customer tenants associated with the partner account and pipe the Customer ID.
    $customerIDs = Get-PartnerCustomer | Select-Object -ExpandProperty CustomerId
    # Loop through each ID and get customer subscriptions filtered for Exchange Online and Microsoft 365 subscriptions.
    foreach ($id in $customerIDs) {
        $subscriptions = Get-PartnerCustomerSubscribedSku -CustomerId $id
        if ($subscriptions.ProductName -match "Exchange Online|Microsoft 365") {
            Write-Host "Gathering information from $($id) ..."
            Write-Log -Message "Gathering information from $($id) ..."
            try {
                $customerTenants = Get-PartnerCustomer -CustomerId $id | Select-Object -ExpandProperty Domain
                $exchangeOnlineTenants += $customerTenants
            }
            catch {
               Write-Host "Get-PartnerCustomer function failed for $($id)."
               Write-Log -Message "Get-PartnerCustomer function failed for $($id)."
            }
        }
        else {
            $customerTenants = Get-PartnerCustomer -CustomerId $id | Select-Object -ExpandProperty Domain
            Write-Host "Gathering information from $($id) ..."
            Write-Host "Skipping '$($customerTenants)' because they do not have any applicable subscriptoins."
            Write-Log -Message "Skipping '$($customerTenants)' because they do not have any applicable subscriptions."
        }
    }

    # Disconnect
    Disconnect-PartnerCenter
    # Return the tenant names
    return $exchangeOnlineTenants
}

# ================================================================================
# Main script body starts here

# Install all required modules
$installModules = "PartnerCenter", "ExchangeOnlineManagement", "AzureAD", "Microsoft.Identity.Client"
Install-RequiredModules -Modules $installModules

# Import Necessary Modules
Import-Module PartnerCenter
Import-Module ExchangeOnlineManagement

# Get user credentials for which the username will be used in the Connect-ExchangeOnline cmdlet later.
$credentials = Get-Credential -Message "Enter your Microsoft Partner username" -ErrorAction Stop
if (-not $credentials) {
    Write-Log -Message "User cancelled Get-Credentials"
    return
}
else {
    $userName = $credentials.UserName
}

# Read the CSV file with the domain safe list
$csvFilePath = Join-Path $PSScriptRoot "Safe Domains.csv"

# Make sure the CSV file exists
if (-not (Test-Path $csvFilePath)) {
    Write-Log -Message "CSV file 'Safe Domains.csv' not found in the script directory."
    Write-Host "CSV file 'Safe Domains.csv' not found in the script directory."
    return
}

# Set the domain list bound for the -SenderDomainIs parameter.
$newSafeDomains = Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty Domain

# Call the Get-PartnerTenants. Only include objects from the partner center that are valid Microsoft tenants.
$tenantDomains = Get-PartnerTenants | Where-Object { $_ -like "*.onmicrosoft.com"}

foreach ($tenantDomain in $tenantDomains) {
    # Establish ExchangeOnline connection
    try {
        Connect-ExchangeOnline -DelegatedOrganization $tenantDomain -ShowProgress $true -UserPrincipalName $userName
    }
    catch {
        Write-Log -Message "$($tenantDomain): Could not connect."
        Write-Host "Could not connect to $($tenantDomain)."
        return
    }
    # Retrieve the existing transport rule, if it exists
    try {
        $existingRule = Get-TransportRule -Identity $ruleName
    }
    catch {
        Write-Log -Message "$($tenantDomain): Get-TransportRule cmdlet failed."
    }

    # If the rule already exists...
    if ($existingRule) {
        # Logging: Existing rule found
        Write-Log -Message "$($tenantDomain): The transport rule '$($ruleName)' already exists. Overwriting with the current safe sender list if it's different."

        # Retrieve the existing domain names from the rule
        $existingSafeDomains = $existingRule.SenderDomainIs

        # Filter out the domain names that are already present in the rule
        $uniqueSafeDomains = $newSafeDomains | Where-Object { $existingSafeDomains -notcontains $_ }

        # If there are names that weren't already in .SenderDomainIs...
        if ($uniqueSafeDomains) {
            # Logging: Domain names to be added
            Write-Log -Message "$($tenantDomain): The rules from 'Safe Domains.CSV' the source file at are different, so we will invoke Set-TransportRule and add the following domains to the safe-list:"
            Write-Log -Message $($newSafeDomains -join ', ')
            Write-Host "$($tenantDomain): The rules from 'Safe Domains.CSV' the source file at are different, so we will invoke Set-TransportRule and add the following domains to the safe-list:"
            Write-Host $($newSafeDomains -join ', ')

            # Set the new rule
            Set-TransportRule -Identity $existingRule.Identity -SenderDomainIs $newSafeDomains

            # Logging: Rule updated
            Write-Log -Message "$($tenantDomain): The transport rule has been updated with the new domain names."
            Write-Host "$($tenantDomain): The transport rule has been updated with the new domain names."
        }
        else {
            # Logging: No new domain names to add
            Write-Log -Message "$($tenantDomain): No new domain names to add to the transport rule."
            Write-Host "$($tenantDomain): No new domain names to add to the transport rule."
        }
    }

    # If the rule doesn't already exist, let's create it!
    elseif (-not $existingrule) {
        # Make sure the CSV file isn't blank or cannot be found
        if (-not $newSafeDomains) {
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
            -SenderDomainIs $newSafeDomains `
            -SetHeaderName X-ETR `
            -SetHeaderValue "Bypass spam filtering for authenticated sender" `
            -SetSCL -1

        Write-Log -Message "$($tenantDomain): The following domains were added to the Firewell Technology Solutions 'Safe Senders' list under the rule named '$ruleName':"
        Write-Log -Message $($newSafeDomains -join ', ')
        Write-Host "$($tenantDomain): The following domains were added to the Firewell Technology Solutions 'Safe Senders' list under the rule named '$ruleName':"
        Write-Output $($newSafeDomains -join ', ')
    }
}

Disconnect-ExchangeOnline -Confirm:$false