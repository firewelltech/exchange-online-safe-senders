# Set the log file path
$logFilePath = "$PSScriptRoot\log.txt"
# Set parameters for Connect-ExchangeOnline cmdlet for efficiency (passed to the -CommandName parameter)

# Check for necessary modules
# AzureAD
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Log -Message "The AzureAD module is not installed."
    try {
    Install-Module -Name AzureAD -force -scope CurrentUser
    }
    catch {
        Write-Log -Message "Failed to install the AzureAD module. Please install it manually and re-run the script."
        Write-Host "Failed to install the AzureAD module. Please install it manually and re-run the script."
        return
    }
    Write-Log -Message "The AzureAD module has been installed"
}
# PartnerCenter
if (-not (Get-Module -ListAvailable -Name PartnerCenter)) {
    Write-Log -Message "The PartnerCenter module is not installed."
    try {
    Install-Module -Name PartnerCenter -force -scope CurrentUser
    }
    catch {
        Write-Log -Message "Failed to install the PartnerCenter module. Please install it manually and re-run the script."
        Write-Host "Failed to install the PartnerCenter module. Please install it manually and re-run the script."
        return
    }
    Write-Log -Message "The PartnerCenter module has been installed"
}
# ExchangeOnlineManagement
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
   try {
     Install-Module -Name ExchangeOnlineManagement -force -scope CurrentUser
   }
   catch {
    Write-Log -Message "Failed to install the ExchangeOnlineManagement module. Please install it manually and re-run the script."
    Write-Host "Failed to install the ExchangeOnlineManagement module. Please install it manually and re-run the script."
    return
   }
}

# Microsoft.Identity.Client
if (-not (Get-Module -ListAvailable -Name Microsoft.Identity.Client)) {
    Write-Log -Message "The Microsoft.Identity.Client module is not installed."
    try {
    Install-Module -Name Microsoft.Identity.Client -force
    }
    catch {
        Write-Log -Message "Failed to install the Microsoft.Identity.Client module. Please install it manually and re-run the script."
        Write-Host "Failed to install the Microsoft.Identity.Client module. Please install it manually and re-run the script."
        return
    }
    Write-Log -Message "The Microsoft.Identity.Client module has been installed"
}

# Function to write log messages to the log file
function Write-Log {
    param (
        [string]$Message
    )

    $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    $logMessage | Out-File -FilePath $logFilePath -Append -Encoding utf8
}

# Get user credentials
$credentials = Get-Credential
$userName = $credentials.UserName

function Get-PartnerTenants {
    $exchangeOnlineTenants = @()
    # Authenticate with Partner Center credentials
    try {
            Connect-PartnerCenter
    }
    catch {
        Write-Log -Message "Failed to connect to Partner Center. Error: $_"
        Write-Host "Failed to connect to Partner Center. Please check your credentials and try again."
        return
    }   
    # Retrieve the customer tenants associated with the partner account and pipe the domain names
    $customerIDs = Get-PartnerCustomer | Select-Object -ExpandProperty CustomerId
    # Get customer subscriptions and filter based on product names because we only want to look at tenants that have Exchange Online or Microsoft 365 subscriptions.
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

# Import Necessary Modules
Import-Module PartnerCenter
Import-Module ExchangeOnlineManagement

# Call the Get-PartnerTenants function to get the domain names of all partner tenants
# Position 0 in the array (the partner tenant itself) should be omitted from the foreach loop later
$tenantDomains = Get-PartnerTenants | Where-Object { $_ -like "*.onmicrosoft.com"}

# Read the CSV file with the domain safe list
$csvFilePath = Join-Path $PSScriptRoot "Safe Domains.csv"
if (-not (Test-Path $csvFilePath)) {
    Write-Log -Message "CSV file 'Safe Domains.csv' not found in the script directory."
    Write-Host "CSV file 'Safe Domains.csv' not found in the script directory."
    return
}

# Set the transport rule name
$ruleName = "Firewell Technology Solutions Safe Senders"

# Set the domain list bound for the -SenderDomainIs parameter.
$newSafeDomains = Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty Domain

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
        $existingRule = Get-TransportRule -Identity $ruleName -ErrorAction SilentlyContinue
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

Disconnect-ExchangeOnline