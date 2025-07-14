<#
.SYNOPSIS
    Comprehensive Exchange Infrastructure Documentation Script
    
.DESCRIPTION
    This script connects to Exchange On-Premises and/or Exchange Online to generate
    comprehensive documentation including ALL configurations, settings, certificates,
    and infrastructure details. Outputs both CSV and HTML reports for auditing and analysis.
    
.PARAMETER Environment
    Specifies the environment to document: OnPremises, Online, or Both
    
.PARAMETER OutputPath
    Specifies the output directory for reports (default: current directory)
    
.PARAMETER ExchangeServer
    For on-premises: FQDN of Exchange server to connect to
    
.PARAMETER Credential
    Credentials for authentication (if not provided, will prompt)
    
.PARAMETER TenantId
    Azure AD Tenant ID for Exchange Online connection (optional)
    
.PARAMETER AppId
    Application ID for certificate-based authentication (optional)
    
.PARAMETER CertificateThumbprint
    Certificate thumbprint for certificate-based authentication (optional)
    
.PARAMETER IncludeDetailedStats
    Include detailed mailbox and database statistics (may take longer)
    
.EXAMPLE
    .\Exchange-Documentation-Script-Enhanced.ps1 -Environment Both -OutputPath "C:\Reports" -IncludeDetailedStats

.EXAMPLE
    .\Exchange-Documentation-Script-Enhanced.ps1 -Environment Both -OutputPath "C:\Reports"
    
.EXAMPLE
    .\Exchange-Documentation-Script-Enhanced.ps1 -Environment OnPremises -ExchangeServer "exchange01.contoso.com"
    
.EXAMPLE
    .\Exchange-Documentation-Script-Enhanced.ps1 -Environment Online -TenantId "contoso.onmicrosoft.com"
    
.EXAMPLE
    .\Exchange-Documentation-Script-Enhanced.ps1 -Environment Online -AppId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABC123..." -TenantId "contoso.onmicrosoft.com"    
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("OnPremises", "Online", "Both")]
    [string]$Environment,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = (Get-Location).Path,
    
    [Parameter(Mandatory=$false)]
    [string]$ExchangeServer,
    
    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$AppId,
    
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDetailedStats
)

# Global variables
$Script:ReportData = @{}
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:CSVPath = Join-Path $OutputPath "Exchange_Comprehensive_Documentation_$Script:Timestamp.csv"
$Script:HTMLPath = Join-Path $OutputPath "Exchange_Comprehensive_Documentation_$Script:Timestamp.html"
$Script:ConnectedToEXO = $false
$Script:ConnectedToGraph = $false

# Function to write progress and log
function Write-LogProgress {
    param([string]$Message, [string]$Status = "Processing")
    Write-Progress -Activity "Exchange Comprehensive Documentation" -Status $Status -CurrentOperation $Message
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message" -ForegroundColor Green
}

# Function to safely execute commands and handle errors
function Invoke-SafeCommand {
    param(
        [scriptblock]$Command,
        [string]$Description,
        [string]$Category
    )
    
    try {
        Write-LogProgress "Collecting $Description"
        $result = & $Command
        if ($result) {
            $Script:ReportData[$Category] = $result
        }
        return $result
    }
    catch {
        Write-Warning "Failed to collect $Description`: $($_.Exception.Message)"
        return $null
    }
}

# Function to check and install required modules
function Test-RequiredModules {
    param([string]$Environment)
    
    $modulesNeeded = @()
    
    if ($Environment -eq "Online" -or $Environment -eq "Both") {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $modulesNeeded += "ExchangeOnlineManagement"
        }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Warning "Microsoft.Graph module not found. Some additional data collection will be skipped."
        }
    }
    
    if ($modulesNeeded.Count -gt 0) {
        Write-Host "The following modules are required but not installed:" -ForegroundColor Yellow
        $modulesNeeded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        Write-Host ""
        Write-Host "To install these modules, run:" -ForegroundColor Cyan
        $modulesNeeded | ForEach-Object { Write-Host "  Install-Module -Name $_ -Scope CurrentUser" -ForegroundColor Cyan }
        Write-Host ""
        
        $install = Read-Host "Would you like to install these modules now? (Y/N)"
        if ($install -eq "Y" -or $install -eq "y") {
            foreach ($module in $modulesNeeded) {
                try {
                    Write-LogProgress "Installing module: $module"
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    Write-Host "Successfully installed $module" -ForegroundColor Green
                }
                catch {
                    Write-Error "Failed to install $module`: $($_.Exception.Message)"
                    return $false
                }
            }
        } else {
            Write-Error "Required modules are not installed. Exiting."
            return $false
        }
    }
    
    return $true
}

# Function to connect to Exchange On-Premises
function Connect-ExchangeOnPremises {
    param([string]$Server, [PSCredential]$Cred)
    
    Write-LogProgress "Connecting to Exchange On-Premises: $Server"
    
    try {
        if ($Cred) {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" -Authentication Kerberos -Credential $Cred
        } else {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/PowerShell/" -Authentication Kerberos
        }
        
        Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
        Write-LogProgress "Successfully connected to Exchange On-Premises"
        return $Session
    }
    catch {
        Write-Error "Failed to connect to Exchange On-Premises: $($_.Exception.Message)"
        return $null
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnline {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$CertThumbprint
    )
    
    Write-LogProgress "Connecting to Exchange Online"
    
    try {
        Import-Module ExchangeOnlineManagement -Force
        
        # Determine connection method
        if ($AppId -and $CertThumbprint -and $TenantId) {
            # Certificate-based authentication
            Write-LogProgress "Using certificate-based authentication"
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertThumbprint -Organization $TenantId -ShowProgress $false
        } elseif ($TenantId) {
            # Interactive authentication with specific tenant
            Write-LogProgress "Using interactive authentication with tenant: $TenantId"
            Connect-ExchangeOnline -UserPrincipalName "admin@$TenantId" -ShowProgress $false
        } else {
            # Standard interactive authentication
            Write-LogProgress "Using interactive authentication"
            Connect-ExchangeOnline -ShowProgress $false
        }
        
        # Test connection
        $null = Get-OrganizationConfig -ErrorAction Stop
        $Script:ConnectedToEXO = $true
        
        Write-LogProgress "Successfully connected to Exchange Online"
        return $true
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        return $false
    }
}

# Function to connect to Microsoft Graph
function Connect-MicrosoftGraph {
    param([string]$TenantId)
    
    try {
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
            Write-LogProgress "Connecting to Microsoft Graph"
            Import-Module Microsoft.Graph.Authentication -Force
            
            $scopes = @(
                "Directory.Read.All",
                "Organization.Read.All",
                "Policy.Read.All",
                "SecurityEvents.Read.All"
            )
            
            if ($TenantId) {
                Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
            } else {
                Connect-MgGraph -Scopes $scopes -NoWelcome
            }
            
            $Script:ConnectedToGraph = $true
            Write-LogProgress "Successfully connected to Microsoft Graph"
            return $true
        }
    }
    catch {
        Write-Warning "Could not connect to Microsoft Graph: $($_.Exception.Message)"
        return $false
    }
}

# Function to collect Exchange On-Premises data
function Get-ExchangeOnPremisesData {
    Write-LogProgress "Starting comprehensive Exchange On-Premises data collection"
    
    # Organization Configuration
    Invoke-SafeCommand -Command {
        Get-OrganizationConfig | Select-Object Name, ExchangeVersion, AdminDisplayVersion, IsDehydrated, 
        HybridConfigurationStatus, MaxReceiveSize, MaxSendSize, DefaultPublicFolderDatabase,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Configuration" -Category "OrganizationConfig"
    
    # Exchange Servers with detailed information
    Invoke-SafeCommand -Command {
        Get-ExchangeServer | Select-Object Name, ServerRole, AdminDisplayVersion, Edition, FQDN, Site, 
        IsHubTransportServer, IsClientAccessServer, IsMailboxServer, IsUnifiedMessagingServer, IsEdgeServer,
        NetworkAddress, OrganizationalUnit, WhenCreated, WhenChanged,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Exchange Servers" -Category "ExchangeServers"
    
    # Exchange Certificates - CRITICAL for SMTP, EWS, etc.
    Invoke-SafeCommand -Command {
        $certs = @()
        $servers = Get-ExchangeServer
        foreach ($server in $servers) {
            try {
                $serverCerts = Get-ExchangeCertificate -Server $server.Name | Select-Object `
                    @{N='Server';E={$server.Name}}, `
                    Thumbprint, Subject, Issuer, NotBefore, NotAfter, Status, `
                    Services, CertificateDomains, IsSelfSigned, HasPrivateKey, `
                    @{N='DaysUntilExpiry';E={($_.NotAfter - (Get-Date)).Days}}, `
                    @{N='IsExpired';E={$_.NotAfter -lt (Get-Date)}}, `
                    @{N='CollectedDate';E={Get-Date}}
                $certs += $serverCerts
            }
            catch {
                Write-Warning "Could not retrieve certificates from server $($server.Name): $($_.Exception.Message)"
            }
        }
        return $certs
    } -Description "Exchange Certificates" -Category "ExchangeCertificates"
    
    # Database Information with detailed settings
    Invoke-SafeCommand -Command {
        Get-MailboxDatabase | Select-Object Name, Server, MasterServerOrAvailabilityGroup, EdbFilePath, 
        LogFolderPath, CircularLoggingEnabled, MaintenanceSchedule, QuotaNotificationSchedule, 
        ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, DeletedItemRetention,
        MailboxRetention, IndexEnabled, BackgroundDatabaseMaintenance, AllowFileRestore,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mailbox Databases" -Category "MailboxDatabases"
    
    # Database Copies and Health
    Invoke-SafeCommand -Command {
        Get-MailboxDatabaseCopyStatus | Select-Object Name, Status, CopyQueueLength, ReplayQueueLength,
        LastInspectedLogTime, ContentIndexState, ActivationSuspended, 
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Database Copy Status" -Category "DatabaseCopyStatus"
    
    # Public Folder Databases (if any)
    Invoke-SafeCommand -Command {
        Get-PublicFolderDatabase -ErrorAction SilentlyContinue | Select-Object Name, Server, EdbFilePath, 
        LogFolderPath, MaintenanceSchedule, MaxItemSize, ProhibitPostQuota, IssueWarningQuota,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Public Folder Databases" -Category "PublicFolderDatabases"
    
    # Database Availability Groups with detailed configuration
    Invoke-SafeCommand -Command {
        Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue | Select-Object Name, Servers, 
        WitnessServer, WitnessDirectory, AlternateWitnessServer, NetworkCompression, NetworkEncryption,
        ReplicationPort, DatacenterActivationMode, ThirdPartyReplication,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Database Availability Groups" -Category "DatabaseAvailabilityGroups"
    
    # Receive Connectors - Including SMTP Relay configurations
    Invoke-SafeCommand -Command {
        Get-ReceiveConnector | Select-Object Identity, Server, Bindings, RemoteIPRanges, AuthMechanism, 
        PermissionGroups, MaxMessageSize, ConnectionTimeout, MaxInboundConnection, RequireTLS,
        EnableAuthGSSAPI, ExtendedProtectionPolicy, SuppressXAnonymousTls, AdvertiseClientSettings,
        Banner, Comment, Enabled, Fqdn, LongAddressesEnabled, OrarEnabled, PipeliningEnabled,
        ProtocolLoggingLevel, SizeEnabled, TarpitInterval, TransportRole,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Receive Connectors (SMTP Relay)" -Category "ReceiveConnectors"
    
    # Send Connectors - Including SMTP Relay configurations
    Invoke-SafeCommand -Command {
        Get-SendConnector | Select-Object Identity, AddressSpaces, SourceTransportServers, SmartHosts, 
        Port, RequireTLS, SmartHostAuthMechanism, UseExternalDNSServersEnabled, MaxMessageSize,
        ConnectionInactivityTimeout, DnsRoutingEnabled, ErrorPolicies, ForceHELO, Fqdn,
        IgnoreSTARTTLS, IsScopedConnector, IsSmtpConnector, LinkedReceiveConnector, ProtocolLoggingLevel,
        SmartHostsString, TlsAuthLevel, TlsCertificateName, TlsDomain,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Send Connectors (SMTP Relay)" -Category "SendConnectors"
    
    # Transport Configuration
    Invoke-SafeCommand -Command {
        Get-TransportConfig | Select-Object MaxDumpsterSizePerDatabase, MaxDumpsterTime, 
        MaxReceiveSize, MaxSendSize, ExternalPostmasterAddress, GenerateCopyOfDSNFor,
        InternalSMTPServers, JournalingReportNdrTo, MaxRecipientEnvelopeLimit, 
        OrganizationFederatedMailbox, RedirectUnprovisionedUserMessagesTo, ShadowRedundancyEnabled,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Configuration" -Category "TransportConfiguration"
    
    # Transport Rules with detailed conditions and actions
    Invoke-SafeCommand -Command {
        Get-TransportRule | Select-Object Name, Priority, State, Mode, Description, Conditions, Actions,
        Exceptions, Comments, RuleVersion, WhenChanged, 
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Rules" -Category "TransportRules"
    
    # Accepted Domains
    Invoke-SafeCommand -Command {
        Get-AcceptedDomain | Select-Object Name, DomainName, DomainType, Default, MatchSubDomains,
        AddressBookEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Accepted Domains" -Category "AcceptedDomains"
    
    # Remote Domains
    Invoke-SafeCommand -Command {
        Get-RemoteDomain | Select-Object Name, DomainName, AllowedOOFType, AutoReplyEnabled,
        AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled,
        UseSimpleDisplayName, @{N='CollectedDate';E={Get-Date}}
    } -Description "Remote Domains" -Category "RemoteDomains"
    
    # Email Address Policies
    Invoke-SafeCommand -Command {
        Get-EmailAddressPolicy | Select-Object Name, Priority, EnabledEmailAddressTemplates, 
        RecipientFilter, RecipientContainer, @{N='CollectedDate';E={Get-Date}}
    } -Description "Email Address Policies" -Category "EmailAddressPolicies"
    
    # Virtual Directories - Critical for client connectivity
    Invoke-SafeCommand -Command {
        $vdirs = @()
        
        # OWA Virtual Directories
        $vdirs += Get-OwaVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'OWA'}}, DefaultDomain, LogonFormat, ClientAuthCleanupLevel, 
            ExternalAuthenticationMethods, InternalAuthenticationMethods, WindowsAuthentication,
            @{N='CollectedDate';E={Get-Date}}
        
        # ECP Virtual Directories
        $vdirs += Get-EcpVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'ECP'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            @{N='CollectedDate';E={Get-Date}}
        
        # ActiveSync Virtual Directories
        $vdirs += Get-ActiveSyncVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'ActiveSync'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            ClientCertAuth, CompressionEnabled, WindowsAuthEnabled,
            @{N='CollectedDate';E={Get-Date}}
        
        # EWS Virtual Directories - CRITICAL
        $vdirs += Get-WebServicesVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'EWS'}}, CertificateAuthentication, WSSecurityAuthentication, OAuthAuthentication,
            ExternalAuthenticationMethods, InternalAuthenticationMethods, WindowsAuthentication,
            @{N='CollectedDate';E={Get-Date}}
        
        # OAB Virtual Directories
        $vdirs += Get-OabVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'OAB'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            RequireSSL, @{N='CollectedDate';E={Get-Date}}
        
        # Autodiscover Virtual Directories
        $vdirs += Get-AutodiscoverVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'Autodiscover'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            WindowsAuthentication, WSSecurityAuthentication,
            @{N='CollectedDate';E={Get-Date}}
        
        # MAPI Virtual Directories
        $vdirs += Get-MapiVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'MAPI'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            @{N='CollectedDate';E={Get-Date}}
        
        # PowerShell Virtual Directories
        $vdirs += Get-PowerShellVirtualDirectory | Select-Object Identity, Server, InternalUrl, ExternalUrl, 
            @{N='Type';E={'PowerShell'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
            RequireSSL, CertificateAuthentication,
            @{N='CollectedDate';E={Get-Date}}
        
        return $vdirs
    } -Description "Virtual Directories (OWA, EWS, ActiveSync, etc.)" -Category "VirtualDirectories"
    
    # Client Access Services
    Invoke-SafeCommand -Command {
        Get-ClientAccessService | Select-Object Name, Server, AutoDiscoverServiceInternalUri, 
        AutoDiscoverSiteScope, AlternateServiceAccountConfiguration,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Client Access Services" -Category "ClientAccessServices"
    
    # Outlook Anywhere Configuration
    Invoke-SafeCommand -Command {
        Get-OutlookAnywhere | Select-Object Identity, Server, InternalHostname, ExternalHostname,
        InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod, IISAuthenticationMethods,
        SSLOffloading, ExternalClientsRequireSsl, InternalClientsRequireSsl,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Outlook Anywhere (RPC over HTTP)" -Category "OutlookAnywhere"
    
    # Federation Trust and Organization Relationships
    Invoke-SafeCommand -Command {
        Get-FederationTrust -ErrorAction SilentlyContinue | Select-Object Name, ApplicationUri, 
        TokenIssuerUris, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Federation Trust" -Category "FederationTrust"
    
    Invoke-SafeCommand -Command {
        Get-OrganizationRelationship -ErrorAction SilentlyContinue | Select-Object Name, DomainNames, 
        FreeBusyAccessEnabled, FreeBusyAccessLevel, FreeBusyAccessScope, MailboxMoveEnabled,
        DeliveryReportEnabled, MailTipsAccessEnabled, MailTipsAccessLevel, MailTipsAccessScope,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Relationships" -Category "OrganizationRelationships"
    
    # Sharing Policies
    Invoke-SafeCommand -Command {
        Get-SharingPolicy | Select-Object Name, Domains, Enabled, Default,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Sharing Policies" -Category "SharingPolicies"
    
    # Retention Policies and Tags
    Invoke-SafeCommand -Command {
        Get-RetentionPolicy | Select-Object Name, RetentionPolicyTagLinks, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policies" -Category "RetentionPolicies"
    
    Invoke-SafeCommand -Command {
        Get-RetentionPolicyTag | Select-Object Name, Type, RetentionEnabled, AgeLimitForRetention,
        RetentionAction, MessageClass, @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policy Tags" -Category "RetentionPolicyTags"
    
    # Address Lists and Global Address Lists
    Invoke-SafeCommand -Command {
        Get-AddressList | Select-Object Name, RecipientFilter, RecipientContainer, DisplayName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Address Lists" -Category "AddressLists"
    
    Invoke-SafeCommand -Command {
        Get-GlobalAddressList | Select-Object Name, RecipientFilter, RecipientContainer,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Global Address Lists" -Category "GlobalAddressLists"
    
    # Offline Address Books
    Invoke-SafeCommand -Command {
        Get-OfflineAddressBook | Select-Object Name, AddressLists, Server, PublicFolderDatabase,
        Schedule, IsDefault, @{N='CollectedDate';E={Get-Date}}
    } -Description "Offline Address Books" -Category "OfflineAddressBooks"
    
    # Mailbox Statistics Summary
    Invoke-SafeCommand -Command {
        $mailboxes = Get-Mailbox -ResultSize Unlimited
        $stats = @{
            TotalMailboxes = $mailboxes.Count
            UserMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}).Count
            SharedMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'}).Count
            ResourceMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -like '*Resource*'}).Count
            RoomMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'RoomMailbox'}).Count
            EquipmentMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'EquipmentMailbox'}).Count
            LinkedMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'LinkedMailbox'}).Count
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$stats
    } -Description "Mailbox Statistics" -Category "MailboxStatistics"
    
    # Detailed Mailbox Statistics (if requested)
    if ($IncludeDetailedStats) {
        Invoke-SafeCommand -Command {
            Write-LogProgress "Collecting detailed mailbox statistics (this will take time)"
            $mailboxStats = @()
            $mailboxes = Get-Mailbox -ResultSize 100 # Limit for performance
            foreach ($mailbox in $mailboxes) {
                try {
                    $stats = Get-MailboxStatistics $mailbox.Identity
                    $mailboxStats += [PSCustomObject]@{
                        DisplayName = $mailbox.DisplayName
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $mailbox.RecipientTypeDetails
                        Database = $stats.Database
                        TotalItemSize = $stats.TotalItemSize
                        ItemCount = $stats.ItemCount
                        DeletedItemCount = $stats.DeletedItemCount
                        LastLogonTime = $stats.LastLogonTime
                        LastLoggedOnUserAccount = $stats.LastLoggedOnUserAccount
                        CollectedDate = Get-Date
                    }
                }
                catch {
                    Write-Warning "Could not get statistics for mailbox $($mailbox.DisplayName)"
                }
            }
            return $mailboxStats
        } -Description "Detailed Mailbox Statistics (Sample)" -Category "DetailedMailboxStats"
    }
    
    # Hybrid Configuration (if exists)
    Invoke-SafeCommand -Command {
        Get-HybridConfiguration -ErrorAction SilentlyContinue | Select-Object Identity, 
        OnPremisesSmartHost, Domains, Features, TlsCertificateName, EdgeTransportServers,
        ReceivingTransportServers, SendingTransportServers, ClientAccessServers,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Hybrid Configuration" -Category "HybridConfiguration"
    
    # Edge Synchronization (if applicable)
    Invoke-SafeCommand -Command {
        Get-EdgeSubscription -ErrorAction SilentlyContinue | Select-Object Name, Site, Domain,
        CreateDate, @{N='CollectedDate';E={Get-Date}}
    } -Description "Edge Subscriptions" -Category "EdgeSubscriptions"
    
    # Message Classifications
    Invoke-SafeCommand -Command {
        Get-MessageClassification -ErrorAction SilentlyContinue | Select-Object Name, DisplayName,
        SenderDescription, RecipientDescription, ClassificationID,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Message Classifications" -Category "MessageClassifications"
    
    # Throttling Policies
    Invoke-SafeCommand -Command {
        Get-ThrottlingPolicy | Select-Object Name, IsDefault, AnonymousMaxConcurrency, 
        EasMaxConcurrency, EwsMaxConcurrency, ImapMaxConcurrency, OutlookServiceMaxConcurrency,
        OwaMaxConcurrency, PopMaxConcurrency, PowerShellMaxConcurrency, RcaMaxConcurrency,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Throttling Policies" -Category "ThrottlingPolicies"
    
    # Mobile Device Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-ActiveSyncMailboxPolicy | Select-Object Name, AllowNonProvisionableDevices, 
        PasswordEnabled, AlphanumericPasswordRequired, PasswordRecoveryEnabled, DeviceEncryptionEnabled,
        AttachmentsEnabled, MaxAttachmentSize, AllowStorageCard, AllowCamera, AllowWiFi, AllowBluetooth,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Mailbox Policies" -Category "MobileDeviceMailboxPolicies"
    
    # OWA Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-OwaMailboxPolicy | Select-Object Name, DirectFileAccessOnPublicComputersEnabled,
        DirectFileAccessOnPrivateComputersEnabled, WebReadyDocumentViewingOnPublicComputersEnabled,
        ForceWebReadyDocumentViewingFirstOnPublicComputers, ActiveSyncIntegrationEnabled,
        AllowOfflineOn, ExternalImageProxyEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "OWA Mailbox Policies" -Category "OWAMailboxPolicies"
    
    # Journal Rules
    Invoke-SafeCommand -Command {
        Get-JournalRule | Select-Object Name, JournalEmailAddress, Scope, Recipient, Enabled,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Journal Rules" -Category "JournalRules"
    
    # Management Role Assignments (Security)
    Invoke-SafeCommand -Command {
        Get-ManagementRoleAssignment | Select-Object Name, Role, RoleAssignee, RoleAssigneeType,
        AssignmentMethod, IsValid, @{N='CollectedDate';E={Get-Date}}
    } -Description "Management Role Assignments" -Category "ManagementRoleAssignments"
    
    Write-LogProgress "Completed comprehensive Exchange On-Premises data collection"
}

# Function to collect Exchange Online data
function Get-ExchangeOnlineData {
    Write-LogProgress "Starting comprehensive Exchange Online data collection"
    
    # Organization Configuration
    Invoke-SafeCommand -Command {
        Get-OrganizationConfig | Select-Object Name, DisplayName, DefaultMailTip, MailTipsAllTipsEnabled,
        MailTipsExternalRecipientsTipsEnabled, MailTipsGroupMetricsEnabled, MailTipsLargeAudienceThreshold,
        IsDehydrated, HybridConfigurationStatus, MaxReceiveSize, MaxSendSize, 
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Configuration" -Category "EXO_OrganizationConfig"
    
    # Tenant Information
    Invoke-SafeCommand -Command {
        $tenant = Get-OrganizationConfig
        $tenantInfo = @{
            TenantName = $tenant.Name
            DisplayName = $tenant.DisplayName
            DefaultDomain = (Get-AcceptedDomain | Where-Object {$_.Default -eq $true}).DomainName
            TotalDomains = (Get-AcceptedDomain).Count
            ExchangeVersion = $tenant.AdminDisplayVersion
            IsDehydrated = $tenant.IsDehydrated
            HybridConfigurationStatus = $tenant.HybridConfigurationStatus
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$tenantInfo
    } -Description "Tenant Information" -Category "EXO_TenantInfo"
    
    # Mailbox Plans
    Invoke-SafeCommand -Command {
        Get-MailboxPlan | Select-Object DisplayName, MaxSendSize, MaxReceiveSize, ProhibitSendQuota,
        ProhibitSendReceiveQuota, IssueWarningQuota, RetainDeletedItemsFor, RoleAssignmentPolicy,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mailbox Plans" -Category "EXO_MailboxPlans"
    
    # Accepted Domains
    Invoke-SafeCommand -Command {
        Get-AcceptedDomain | Select-Object Name, DomainName, DomainType, Default, MatchSubDomains,
        AddressBookEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Accepted Domains" -Category "EXO_AcceptedDomains"
    
    # Remote Domains
    Invoke-SafeCommand -Command {
        Get-RemoteDomain | Select-Object Name, DomainName, AllowedOOFType, AutoReplyEnabled,
        AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled,
        UseSimpleDisplayName, @{N='CollectedDate';E={Get-Date}}
    } -Description "Remote Domains" -Category "EXO_RemoteDomains"
    
    # Transport Configuration
    Invoke-SafeCommand -Command {
        Get-TransportConfig | Select-Object MaxReceiveSize, MaxSendSize, ExternalPostmasterAddress,
        GenerateCopyOfDSNFor, JournalingReportNdrTo, MaxRecipientEnvelopeLimit, 
        OrganizationFederatedMailbox, RedirectUnprovisionedUserMessagesTo,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Configuration" -Category "EXO_TransportConfiguration"
    
    # Transport Rules
    Invoke-SafeCommand -Command {
        Get-TransportRule | Select-Object Name, Priority, State, Mode, Description, Comments,
        RuleVersion, WhenChanged, @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Rules" -Category "EXO_TransportRules"
    
    # Inbound Connectors - SMTP Relay for Exchange Online
    Invoke-SafeCommand -Command {
        Get-InboundConnector | Select-Object Name, ConnectorType, ConnectorSource, SenderDomains,
        SenderIPAddresses, RequireTls, RestrictDomainsToIPAddresses, Enabled, Comment,
        TlsSenderCertificateName, CloudServicesMailEnabled, TreatMessagesAsInternal,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Inbound Connectors (SMTP Relay)" -Category "EXO_InboundConnectors"
    
    # Outbound Connectors - SMTP Relay for Exchange Online
    Invoke-SafeCommand -Command {
        Get-OutboundConnector | Select-Object Name, ConnectorType, RecipientDomains, SmartHosts,
        TlsDomain, UseMxRecord, RouteAllMessagesViaOnPremises, Enabled, Comment,
        TlsSettings, IsTransportRuleScoped, CloudServicesMailEnabled, AllAcceptedDomains,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Outbound Connectors (SMTP Relay)" -Category "EXO_OutboundConnectors"
    
    # Anti-Spam Policies (Exchange Online Protection)
    Invoke-SafeCommand -Command {
        Get-HostedContentFilterPolicy | Select-Object Name, SpamAction, HighConfidenceSpamAction,
        PhishSpamAction, BulkSpamAction, QuarantineRetentionPeriod, EndUserSpamNotificationFrequency,
        TestModeAction, IncreaseScoreWithImageLinks, IncreaseScoreWithNumericIps,
        IncreaseScoreWithRedirectToOtherPort, IncreaseScoreWithBizOrInfoUrls,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Spam Policies (EOP)" -Category "EXO_AntiSpamPolicies"
    
    # Anti-Malware Policies (Exchange Online Protection)
    Invoke-SafeCommand -Command {
        Get-MalwareFilterPolicy | Select-Object Name, Action, EnableFileFilter, FileTypes,
        EnableInternalSenderAdminNotifications, EnableInternalSenderNotifications,
        InternalSenderAdminAddress, CustomNotifications, CustomFromAddress, CustomFromName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Malware Policies (EOP)" -Category "EXO_AntiMalwarePolicies"
    
    # Connection Filter Policies (IP Allow/Block Lists)
    Invoke-SafeCommand -Command {
        Get-HostedConnectionFilterPolicy | Select-Object Name, IPAllowList, IPBlockList, 
        EnableSafeList, DirectoryBasedEdgeBlockMode, @{N='CollectedDate';E={Get-Date}}
    } -Description "Connection Filter Policies (IP Lists)" -Category "EXO_ConnectionFilterPolicies"
    
    # Safe Attachments Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue | Select-Object Name, Enable, Action,
        Redirect, RedirectAddress, ActionOnError, @{N='CollectedDate';E={Get-Date}}
    } -Description "Safe Attachments Policies (Defender)" -Category "EXO_SafeAttachmentPolicies"
    
    # Safe Links Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-SafeLinksPolicy -ErrorAction SilentlyContinue | Select-Object Name, IsEnabled,
        ScanUrls, EnableForInternalSenders, TrackClicks, AllowClickThrough, EnableSafeLinksForTeams,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Safe Links Policies (Defender)" -Category "EXO_SafeLinksPolicies"
    
    # ATP Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-AtpPolicyForO365 -ErrorAction SilentlyContinue | Select-Object Name, EnableATPForSPOTeamsODB,
        EnableSafeDocs, AllowSafeDocsOpen, @{N='CollectedDate';E={Get-Date}}
    } -Description "ATP Policies (Defender)" -Category "EXO_ATPPolicies"
    
    # Anti-Phishing Policies
    Invoke-SafeCommand -Command {
        Get-AntiPhishPolicy -ErrorAction SilentlyContinue | Select-Object Name, Enabled, 
        EnableMailboxIntelligence, EnableMailboxIntelligenceProtection, EnableSpoofIntelligence,
        EnableFirstContactSafetyTips, EnableSimilarUsersSafetyTips, EnableSimilarDomainsSafetyTips,
        EnableUnusualCharactersSafetyTips, @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Phishing Policies" -Category "EXO_AntiPhishingPolicies"
    
    # DKIM Configuration
    Invoke-SafeCommand -Command {
        $domains = Get-AcceptedDomain | Where-Object {$_.DomainType -ne "InternalRelay"}
        $dkimConfig = @()
        foreach ($domain in $domains) {
            try {
                $dkim = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction SilentlyContinue
                if ($dkim) {
                    $dkimConfig += $dkim | Select-Object Domain, Enabled, Status, Selector1CNAME, 
                        Selector2CNAME, @{N='CollectedDate';E={Get-Date}}
                }
            }
            catch {
                # Domain might not support DKIM
            }
        }
        return $dkimConfig
    } -Description "DKIM Configuration" -Category "EXO_DKIMConfiguration"
    
    # Detailed Mailbox Statistics
    Invoke-SafeCommand -Command {
        Write-LogProgress "Collecting detailed mailbox statistics (this may take a while)"
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum
        $stats = @{
            TotalMailboxes = $mailboxes.Count
            UserMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}).Count
            SharedMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'}).Count
            ResourceMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -like '*Resource*'}).Count
            EquipmentMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'EquipmentMailbox'}).Count
            RoomMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'RoomMailbox'}).Count
            DiscoveryMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'DiscoveryMailbox'}).Count
            GroupMailboxes = ($mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'GroupMailbox'}).Count
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$stats
    } -Description "Mailbox Statistics" -Category "EXO_MailboxStatistics"
    
    # Distribution Groups Statistics
    Invoke-SafeCommand -Command {
        $groups = Get-DistributionGroup -ResultSize Unlimited
        $groupStats = @{
            TotalDistributionGroups = $groups.Count
            SecurityGroups = ($groups | Where-Object {$_.GroupType -like '*Security*'}).Count
            UniversalGroups = ($groups | Where-Object {$_.GroupType -like '*Universal*'}).Count
            DynamicGroups = (Get-DynamicDistributionGroup -ResultSize Unlimited).Count
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$groupStats
    } -Description "Distribution Group Statistics" -Category "EXO_DistributionGroupStats"
    
    # Mobile Device Access Rules
    Invoke-SafeCommand -Command {
        Get-ActiveSyncDeviceAccessRule | Select-Object Identity, Characteristic, QueryString,
        AccessLevel, @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Access Rules" -Category "EXO_MobileDeviceRules"
    
    # Mobile Device Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-ActiveSyncMailboxPolicy | Select-Object Name, AllowNonProvisionableDevices, 
        PasswordEnabled, AlphanumericPasswordRequired, PasswordRecoveryEnabled, DeviceEncryptionEnabled,
        AttachmentsEnabled, MaxAttachmentSize, AllowStorageCard, AllowCamera, AllowWiFi, AllowBluetooth,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Mailbox Policies" -Category "EXO_MobileDevicePolicies"
    
    # OWA Policies
    Invoke-SafeCommand -Command {
        Get-OwaMailboxPolicy | Select-Object Name, DirectFileAccessOnPublicComputersEnabled,
        DirectFileAccessOnPrivateComputersEnabled, WebReadyDocumentViewingOnPublicComputersEnabled,
        ForceWebReadyDocumentViewingFirstOnPublicComputers, ActiveSyncIntegrationEnabled,
        AllowOfflineOn, ExternalImageProxyEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "OWA Policies" -Category "EXO_OWAPolicies"
    
    # Retention Policies
    Invoke-SafeCommand -Command {
        Get-RetentionPolicy | Select-Object Name, RetentionPolicyTagLinks, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policies" -Category "EXO_RetentionPolicies"
    
    # Data Loss Prevention Policies
    Invoke-SafeCommand -Command {
        Get-DlpPolicy -ErrorAction SilentlyContinue | Select-Object Name, State, Mode, Description,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "DLP Policies" -Category "EXO_DLPPolicies"
    
    # Quarantine Policies
    Invoke-SafeCommand -Command {
        Get-QuarantinePolicy -ErrorAction SilentlyContinue | Select-Object Name, 
        EndUserQuarantinePermissionsValue, ESNEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Quarantine Policies" -Category "EXO_QuarantinePolicies"
    
    # Address Lists
    Invoke-SafeCommand -Command {
        Get-AddressList | Select-Object Name, RecipientFilter, DisplayName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Address Lists" -Category "EXO_AddressLists"
    
    # Global Address Lists
    Invoke-SafeCommand -Command {
        Get-GlobalAddressList | Select-Object Name, RecipientFilter,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Global Address Lists" -Category "EXO_GlobalAddressLists"
    
    # Offline Address Books
    Invoke-SafeCommand -Command {
        Get-OfflineAddressBook | Select-Object Name, AddressLists, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Offline Address Books" -Category "EXO_OfflineAddressBooks"
    
    # Organization Relationships (Federation)
    Invoke-SafeCommand -Command {
        Get-OrganizationRelationship | Select-Object Name, DomainNames, 
        FreeBusyAccessEnabled, FreeBusyAccessLevel, FreeBusyAccessScope, MailboxMoveEnabled,
        DeliveryReportEnabled, MailTipsAccessEnabled, MailTipsAccessLevel, MailTipsAccessScope,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Relationships" -Category "EXO_OrganizationRelationships"
    
    # Sharing Policies
    Invoke-SafeCommand -Command {
        Get-SharingPolicy | Select-Object Name, Domains, Enabled, Default,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Sharing Policies" -Category "EXO_SharingPolicies"
    
    # Role Assignment Policies
    Invoke-SafeCommand -Command {
        Get-RoleAssignmentPolicy | Select-Object Name, Description, IsDefault, AssignedRoles,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Role Assignment Policies" -Category "EXO_RoleAssignmentPolicies"
    
    # Audit Configuration
    Invoke-SafeCommand -Command {
        Get-AdminAuditLogConfig | Select-Object AdminAuditLogEnabled, AdminAuditLogCmdlets,
        AdminAuditLogParameters, AdminAuditLogExcludedCmdlets, LogLevel,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Admin Audit Log Configuration" -Category "EXO_AdminAuditConfig"
    
    # Microsoft Graph data (if connected)
    if ($Script:ConnectedToGraph) {
        Invoke-SafeCommand -Command {
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force
            $org = Get-MgOrganization
            $orgInfo = @{
                DisplayName = $org.DisplayName
                TenantType = $org.TenantType
                CountryLetterCode = $org.CountryLetterCode
                CreatedDateTime = $org.CreatedDateTime
                TechnicalNotificationMails = $org.TechnicalNotificationMails -join ", "
                CollectedDate = Get-Date
            }
            return [PSCustomObject]$orgInfo
        } -Description "Azure AD Organization Info" -Category "AAD_OrganizationInfo"
    }
    
    Write-LogProgress "Completed comprehensive Exchange Online data collection"
}

# Function to export data to CSV
function Export-ToCSV {
    Write-LogProgress "Exporting data to CSV format"
    
    $csvData = @()
    
    foreach ($category in $Script:ReportData.Keys) {
        $data = $Script:ReportData[$category]
        
        if ($data -is [Array]) {
            foreach ($item in $data) {
                $csvRow = [PSCustomObject]@{
                    Category = $category
                    Data = ($item | ConvertTo-Json -Compress)
                    CollectedDate = Get-Date
                }
                $csvData += $csvRow
            }
        } else {
            $csvRow = [PSCustomObject]@{
                Category = $category
                Data = ($data | ConvertTo-Json -Compress)
                CollectedDate = Get-Date
            }
            $csvData += $csvRow
        }
    }
    
    $csvData | Export-Csv -Path $Script:CSVPath -NoTypeInformation -Encoding UTF8
    Write-LogProgress "CSV report saved to: $Script:CSVPath"
}

# Function to generate HTML report
function Export-ToHTML {
    Write-LogProgress "Generating comprehensive HTML report"
    
    # Determine environment type for report title
    $envType = switch ($Environment) {
        "OnPremises" { "On-Premises Exchange" }
        "Online" { "Exchange Online" }
        "Both" { "Hybrid Exchange Environment" }
    }
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$envType Comprehensive Infrastructure Documentation</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; text-align: center; margin-bottom: 30px; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 30px; margin-bottom: 15px; padding: 10px; background-color: #ecf0f1; border-left: 5px solid #3498db; }
        h3 { color: #2c3e50; margin-top: 20px; margin-bottom: 10px; }
        .info-box { background-color: #e8f4fd; border: 1px solid #bee5eb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .warning-box { background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .success-box { background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .critical-box { background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .online-box { background-color: #e1f5fe; border: 1px solid #81d4fa; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .onprem-box { background-color: #f3e5f5; border: 1px solid #ce93d8; border-radius: 5px; padding: 15px; margin: 10px 0; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; background-color: white; font-size: 0.9em; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #3498db; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f8f9fa; }
        tr:hover { background-color: #e8f4fd; }
        .timestamp { color: #7f8c8d; font-style: italic; text-align: center; margin-top: 30px; }
        .summary-stats { display: flex; justify-content: space-around; margin: 20px 0; flex-wrap: wrap; }
        .stat-box { background-color: #3498db; color: white; padding: 20px; border-radius: 10px; text-align: center; min-width: 150px; margin: 5px; }
        .stat-box.online { background-color: #2196f3; }
        .stat-box.onprem { background-color: #9c27b0; }
        .stat-box.critical { background-color: #e74c3c; }
        .stat-box.warning { background-color: #f39c12; }
        .stat-number { font-size: 2em; font-weight: bold; }
        .stat-label { font-size: 0.9em; margin-top: 5px; }
        .no-data { color: #7f8c8d; font-style: italic; text-align: center; padding: 20px; }
        .collapsible { background-color: #3498db; color: white; cursor: pointer; padding: 15px; width: 100%; border: none; text-align: left; outline: none; font-size: 16px; margin-top: 10px; }
        .collapsible:hover { background-color: #2980b9; }
        .collapsible.online { background-color: #2196f3; }
        .collapsible.online:hover { background-color: #1976d2; }
        .collapsible.onprem { background-color: #9c27b0; }
        .collapsible.onprem:hover { background-color: #7b1fa2; }
        .collapsible.critical { background-color: #e74c3c; }
        .collapsible.critical:hover { background-color: #c0392b; }
        .content { padding: 0; display: none; overflow: hidden; background-color: #f8f9fa; }
        .content.show { display: block; padding: 15px; }
        .environment-badge { display: inline-block; padding: 5px 10px; border-radius: 15px; font-size: 0.8em; font-weight: bold; margin-left: 10px; }
        .badge-online { background-color: #2196f3; color: white; }
        .badge-onprem { background-color: #9c27b0; color: white; }
        .badge-critical { background-color: #e74c3c; color: white; }
        .cert-expired { background-color: #ffebee; color: #c62828; }
        .cert-expiring { background-color: #fff3e0; color: #ef6c00; }
        .cert-valid { background-color: #e8f5e8; color: #2e7d32; }
    </style>
    <script>
        function toggleContent(element) {
            var content = element.nextElementSibling;
            content.classList.toggle('show');
            element.textContent = content.classList.contains('show') ? 
                element.textContent.replace('▶', '▼') : 
                element.textContent.replace('▼', '▶');
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>$envType Comprehensive Infrastructure Documentation</h1>
        <div class="info-box">
            <strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>
            <strong>Environment:</strong> $Environment<br>
            <strong>Total Categories:</strong> $($Script:ReportData.Keys.Count)<br>
            <strong>Exchange Online Connected:</strong> $Script:ConnectedToEXO<br>
            <strong>Microsoft Graph Connected:</strong> $Script:ConnectedToGraph<br>
            <strong>Detailed Statistics Included:</strong> $IncludeDetailedStats
        </div>
"@

    # Add critical alerts section
    $criticalAlerts = @()
    
    # Check for expired certificates
    if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
        $expiredCerts = $Script:ReportData["ExchangeCertificates"] | Where-Object {$_.IsExpired -eq $true}
        $expiringSoonCerts = $Script:ReportData["ExchangeCertificates"] | Where-Object {$_.DaysUntilExpiry -le 30 -and $_.DaysUntilExpiry -gt 0}
        
        if ($expiredCerts.Count -gt 0) {
            $criticalAlerts += "⚠️ $($expiredCerts.Count) expired certificate(s) found"
        }
        if ($expiringSoonCerts.Count -gt 0) {
            $criticalAlerts += "⚠️ $($expiringSoonCerts.Count) certificate(s) expiring within 30 days"
        }
    }
    
    if ($criticalAlerts.Count -gt 0) {
        $htmlContent += "<div class='critical-box'><h3>🚨 Critical Alerts</h3><ul>"
        foreach ($alert in $criticalAlerts) {
            $htmlContent += "<li>$alert</li>"
        }
        $htmlContent += "</ul></div>"
    }

    # Add summary statistics if available
    $hasStats = $Script:ReportData.ContainsKey("MailboxStatistics") -or $Script:ReportData.ContainsKey("EXO_MailboxStatistics")
    if ($hasStats) {
        $htmlContent += "<h2>📊 Summary Statistics</h2><div class='summary-stats'>"
        
        if ($Script:ReportData.ContainsKey("MailboxStatistics")) {
            $stats = $Script:ReportData["MailboxStatistics"]
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$($stats.TotalMailboxes)</div><div class='stat-label'>Total Mailboxes (On-Prem)</div></div>"
        }
        
        if ($Script:ReportData.ContainsKey("EXO_MailboxStatistics")) {
            $stats = $Script:ReportData["EXO_MailboxStatistics"]
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$($stats.TotalMailboxes)</div><div class='stat-label'>Total Mailboxes (Online)</div></div>"
        }
        
        if ($Script:ReportData.ContainsKey("ExchangeServers")) {
            $serverCount = $Script:ReportData["ExchangeServers"].Count
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$serverCount</div><div class='stat-label'>Exchange Servers</div></div>"
        }
        
        if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
            $certCount = $Script:ReportData["ExchangeCertificates"].Count
            $expiredCount = ($Script:ReportData["ExchangeCertificates"] | Where-Object {$_.IsExpired -eq $true}).Count
            if ($expiredCount -gt 0) {
                $htmlContent += "<div class='stat-box critical'><div class='stat-number'>$expiredCount</div><div class='stat-label'>Expired Certificates</div></div>"
            }
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$certCount</div><div class='stat-label'>Total Certificates</div></div>"
        }
        
        if ($Script:ReportData.ContainsKey("EXO_AcceptedDomains")) {
            $domainCount = $Script:ReportData["EXO_AcceptedDomains"].Count
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$domainCount</div><div class='stat-label'>Accepted Domains (Online)</div></div>"
        }
        
        $htmlContent += "</div>"
    }

    # Generate sections for each category with enhanced styling
    foreach ($category in $Script:ReportData.Keys | Sort-Object) {
        $data = $Script:ReportData[$category]
        $displayName = $category -replace "_", " " -replace "EXO", "Exchange Online"
        
        # Determine environment type and criticality for styling
        $envClass = "onprem"
        $envBadge = "<span class='environment-badge badge-onprem'>On-Premises</span>"
        
        if ($category.StartsWith("EXO_")) { 
            $envClass = "online"
            $envBadge = "<span class='environment-badge badge-online'>Exchange Online</span>"
        }
        
        # Mark critical categories
        if ($category -eq "ExchangeCertificates" -and $criticalAlerts.Count -gt 0) {
            $envClass = "critical"
            $envBadge += "<span class='environment-badge badge-critical'>Critical</span>"
        }
        
        $htmlContent += "<button class='collapsible $envClass' onclick='toggleContent(this)'>▶ $displayName $envBadge</button>"
        $htmlContent += "<div class='content'>"
        
        if ($data -and (($data -is [Array] -and $data.Count -gt 0) -or ($data -isnot [Array]))) {
            if ($data -is [Array] -and $data[0] -is [PSCustomObject]) {
                # Create table for structured data
                $properties = $data[0].PSObject.Properties.Name
                $htmlContent += "<table><thead><tr>"
                foreach ($prop in $properties) {
                    $htmlContent += "<th>$prop</th>"
                }
                $htmlContent += "</tr></thead><tbody>"
                
                foreach ($item in $data) {
                    $rowClass = ""
                    
                    # Special formatting for certificates
                    if ($category -eq "ExchangeCertificates") {
                        if ($item.IsExpired -eq $true) {
                            $rowClass = "cert-expired"
                        } elseif ($item.DaysUntilExpiry -le 30 -and $item.DaysUntilExpiry -gt 0) {
                            $rowClass = "cert-expiring"
                        } else {
                            $rowClass = "cert-valid"
                        }
                    }
                    
                    $htmlContent += "<tr class='$rowClass'>"
                    foreach ($prop in $properties) {
                        $value = $item.$prop
                        if ($value -is [Array]) {
                            $value = $value -join ", "
                        }
                        $htmlContent += "<td>$([System.Web.HttpUtility]::HtmlEncode($value))</td>"
                    }
                    $htmlContent += "</tr>"
                }
                $htmlContent += "</tbody></table>"
            } elseif ($data -is [PSCustomObject]) {
                # Single object - display as key-value pairs
                $htmlContent += "<table><thead><tr><th>Property</th><th>Value</th></tr></thead><tbody>"
                foreach ($prop in $data.PSObject.Properties) {
                    $value = $prop.Value
                    if ($value -is [Array]) {
                        $value = $value -join ", "
                    }
                    $htmlContent += "<tr><td><strong>$($prop.Name)</strong></td><td>$([System.Web.HttpUtility]::HtmlEncode($value))</td></tr>"
                }
                $htmlContent += "</tbody></table>"
            } else {
                $htmlContent += "<pre>$([System.Web.HttpUtility]::HtmlEncode(($data | Out-String)))</pre>"
            }
        } else {
            $htmlContent += "<div class='no-data'>No data available for this category</div>"
        }
        
        $htmlContent += "</div>"
    }

    $htmlContent += @"
        <div class="timestamp">
            Comprehensive report generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Exchange Documentation Script v3.0<br>
            This report includes SMTP relays, EWS certificates, and all critical Exchange configurations
        </div>
    </div>
</body>
</html>
"@

    # Add System.Web assembly for HTML encoding
    Add-Type -AssemblyName System.Web
    
    $htmlContent | Out-File -FilePath $Script:HTMLPath -Encoding UTF8
    Write-LogProgress "Comprehensive HTML report saved to: $Script:HTMLPath"
}

# Main execution function
function Start-ExchangeDocumentation {
    Write-LogProgress "Starting Comprehensive Exchange Infrastructure Documentation" "Initializing"
    
    # Check required modules
    if (-not (Test-RequiredModules -Environment $Environment)) {
        return
    }
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    $onPremSession = $null
    
    try {
        # Connect and collect data based on environment
        switch ($Environment) {
            "OnPremises" {
                if (-not $ExchangeServer) {
                    $ExchangeServer = Read-Host "Enter Exchange Server FQDN"
                }
                $onPremSession = Connect-ExchangeOnPremises -Server $ExchangeServer -Cred $Credential
                if ($onPremSession) {
                    Get-ExchangeOnPremisesData
                }
            }
            "Online" {
                if (Connect-ExchangeOnline -TenantId $TenantId -AppId $AppId -CertThumbprint $CertificateThumbprint) {
                    # Also try to connect to Microsoft Graph for additional data
                    Connect-MicrosoftGraph -TenantId $TenantId | Out-Null
                    Get-ExchangeOnlineData
                }
            }
            "Both" {
                # On-Premises first
                if (-not $ExchangeServer) {
                    $ExchangeServer = Read-Host "Enter Exchange Server FQDN (or press Enter to skip On-Premises)"
                }
                if ($ExchangeServer) {
                    $onPremSession = Connect-ExchangeOnPremises -Server $ExchangeServer -Cred $Credential
                    if ($onPremSession) {
                        Get-ExchangeOnPremisesData
                    }
                }
                
                # Then Exchange Online
                if (Connect-ExchangeOnline -TenantId $TenantId -AppId $AppId -CertThumbprint $CertificateThumbprint) {
                    Connect-MicrosoftGraph -TenantId $TenantId | Out-Null
                    Get-ExchangeOnlineData
                }
            }
        }
        
        # Generate reports
        if ($Script:ReportData.Keys.Count -gt 0) {
            Export-ToCSV
            Export-ToHTML
            
            Write-Host "`n" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "COMPREHENSIVE EXCHANGE DOCUMENTATION COMPLETED" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "Environment: $Environment" -ForegroundColor Cyan
            Write-Host "CSV Report: $Script:CSVPath" -ForegroundColor Yellow
            Write-Host "HTML Report: $Script:HTMLPath" -ForegroundColor Yellow
            Write-Host "Total Categories Documented: $($Script:ReportData.Keys.Count)" -ForegroundColor Cyan
            Write-Host "Exchange Online Connected: $Script:ConnectedToEXO" -ForegroundColor Cyan
            Write-Host "Microsoft Graph Connected: $Script:ConnectedToGraph" -ForegroundColor Cyan
            Write-Host "Detailed Statistics: $IncludeDetailedStats" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Green
            
            # Show critical alerts in console
            if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
                $expiredCerts = $Script:ReportData["ExchangeCertificates"] | Where-Object {$_.IsExpired -eq $true}
                $expiringSoonCerts = $Script:ReportData["ExchangeCertificates"] | Where-Object {$_.DaysUntilExpiry -le 30 -and $_.DaysUntilExpiry -gt 0}
                
                if ($expiredCerts.Count -gt 0) {
                    Write-Host "⚠️  CRITICAL: $($expiredCerts.Count) expired certificate(s) found!" -ForegroundColor Red
                }
                if ($expiringSoonCerts.Count -gt 0) {
                    Write-Host "⚠️  WARNING: $($expiringSoonCerts.Count) certificate(s) expiring within 30 days!" -ForegroundColor Yellow
                }
            }
            
        } else {
            Write-Warning "No data was collected. Please check your connections and permissions."
        }
    }
    catch {
        Write-Error "An error occurred during documentation: $($_.Exception.Message)"
    }
    finally {
        # Cleanup connections
        if ($onPremSession) {
            Remove-PSSession $onPremSession -ErrorAction SilentlyContinue
        }
        
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        } catch {}
        
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        } catch {}
        
        Write-Progress -Activity "Exchange Comprehensive Documentation" -Completed
    }
}

# Execute the main function
Start-ExchangeDocumentation
