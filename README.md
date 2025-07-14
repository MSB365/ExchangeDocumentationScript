# Comprehensive Exchange Infrastructure Documentation Script

A comprehensive PowerShell script for documenting Microsoft Exchange Server environments, including on-premises Exchange, Exchange Online, and hybrid configurations. This script generates detailed CSV and HTML reports suitable for infrastructure documentation, compliance auditing, and analysis.

## 🚀 Features

### Supported Environments
- **On-Premises Exchange Server** (2013, 2016, 2019)
- **Exchange Online** (Microsoft 365)
- **Hybrid Exchange** environments
- **Exchange Online Protection** (EOP) settings

### 🔧 Critical Infrastructure Components
- **SMTP Relay Configuration**: Complete send/receive connector documentation
- **Exchange Web Services (EWS)**: Virtual directory configurations and certificates
- **Certificate Management**: All Exchange certificates with expiration tracking
- **Transport Security**: TLS settings, authentication methods, and security policies
- **Client Connectivity**: All virtual directories (OWA, ECP, ActiveSync, EWS, OAB, Autodiscover, MAPI, PowerShell)
- **Federation and Hybrid**: Complete hybrid configuration and organization relationships

### Documentation Coverage

#### On-Premises Exchange
- Exchange server inventory and roles
- Database configurations (mailbox, public folder)
- Database Availability Groups (DAG)
- Transport connectors (send/receive)
- Virtual directories (OWA, ECP, ActiveSync, etc.)
- Transport rules and policies
- Accepted domains and email address policies
- Hybrid configuration details
- Client access services
- Mailbox statistics and distribution

#### Additional On-Premises Components:
- **Exchange Certificates**: All certificates with expiration dates and services
- **SMTP Relay Configuration**: Complete send/receive connector settings
- **EWS and Client Access**: All virtual directories with authentication methods
- **Transport Security**: TLS certificates, authentication mechanisms
- **Federation Configuration**: Organization relationships and sharing policies
- **Retention and Compliance**: Retention policies, tags, and journal rules
- **Address Lists and OABs**: Complete directory service configuration
- **Throttling and Mobile Policies**: Client access and device management
- **Management Roles**: Security and administrative role assignments

#### Exchange Online
- Organization configuration
- Tenant information and accepted domains
- Mailbox plans and policies
- Transport rules and connectors
- Exchange Online Protection settings
- Anti-spam and anti-malware policies
- Microsoft Defender for Office 365 policies
- Mobile device policies
- Data Loss Prevention (DLP) policies
- Retention and quarantine policies
- Detailed mailbox and group statistics

#### Additional Exchange Online Components:
- **SMTP Relay for Cloud**: Inbound/outbound connectors for hybrid scenarios
- **DKIM Configuration**: Domain-based message authentication
- **Advanced Threat Protection**: Complete Defender for Office 365 settings
- **Connection Filtering**: IP allow/block lists and connection policies
- **Anti-Phishing Policies**: Advanced phishing protection settings
- **Audit Configuration**: Admin audit logging and compliance settings

### Output Formats
- **CSV Report**: Machine-readable format for data analysis
- **HTML Report**: Interactive, professional report with collapsible sections
- **Comprehensive Statistics**: Summary dashboards and detailed breakdowns

## 🚨 Critical Monitoring Features
- **Certificate Expiration Alerts**: Automatic detection of expired and expiring certificates
- **Security Configuration Review**: Complete authentication and TLS settings
- **SMTP Relay Documentation**: All inbound/outbound connectors with security settings
- **Transport Rule Analysis**: Complete mail flow rule documentation

## 📋 Prerequisites

### PowerShell Modules
The script will automatically check for and optionally install required modules:

#### For Exchange Online
\`\`\`powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
\`\`\`

#### For Microsoft Graph (Optional - provides additional Azure AD data)
\`\`\`powershell
Install-Module -Name Microsoft.Graph -Scope CurrentUser
\`\`\`

### Permissions Required

#### On-Premises Exchange
- Exchange Organization Management role
- Local administrator rights on Exchange server (for PowerShell remoting)

#### Exchange Online
- Exchange Administrator role
- Global Administrator role (for full feature access)
- Security Administrator role (for Defender for Office 365 features)

## 🛠️ Installation

1. **Download the script**
   \`\`\`bash
   git clone https://github.com/yourusername/exchange-documentation-script.git
   cd exchange-documentation-script
   \`\`\`

2. **Set execution policy** (if needed)
   \`\`\`powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   \`\`\`

3. **Install required modules** (script will prompt if needed)
   \`\`\`powershell
   Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
   Install-Module -Name Microsoft.Graph -Scope CurrentUser
   \`\`\`

## 📖 Usage

### Basic Usage Examples

#### Document Exchange Online Only
\`\`\`powershell
.\\Exchange-Documentation-Script.ps1 -Environment Online -OutputPath "C:\\Reports"
\`\`\`

#### Document On-Premises Exchange Only
\`\`\`powershell
.\\Exchange-Documentation-Script.ps1 -Environment OnPremises -ExchangeServer "exchange01.contoso.com" -OutputPath "C:\\Reports"
\`\`\`

#### Document Both Environments (Hybrid)
\`\`\`powershell
.\\Exchange-Documentation-Script.ps1 -Environment Both -ExchangeServer "exchange01.contoso.com" -OutputPath "C:\\Reports"
\`\`\`

### Advanced Usage Examples

#### Exchange Online with Specific Tenant
\`\`\`powershell
.\\Exchange-Documentation-Script.ps1 -Environment Online -TenantId "contoso.onmicrosoft.com" -OutputPath "C:\\Reports"
\`\`\`

#### Certificate-Based Authentication (Exchange Online)
\`\`\`powershell
.\\Exchange-Documentation-Script.ps1 -Environment Online -AppId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABC123DEF456..." -TenantId "contoso.onmicrosoft.com"
\`\`\`

#### On-Premises with Specific Credentials
\`\`\`powershell
\$cred = Get-Credential
.\\Exchange-Documentation-Script.ps1 -Environment OnPremises -ExchangeServer "exchange01.contoso.com" -Credential \$cred
\`\`\`

#### Comprehensive documentation with detailed statistics
\`\`\`powershell
.\\Exchange-Documentation-Script-Enhanced.ps1 -Environment Both -OutputPath "C:\\Reports" -IncludeDetailedStats
\`\`\`

#### Focus on certificate and security analysis
\`\`\`powershell
.\\Exchange-Documentation-Script-Enhanced.ps1 -Environment OnPremises -ExchangeServer "exchange01.contoso.com" -IncludeDetailedStats
\`\`\`

## 📊 Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| \`Environment\` | String | Yes | Environment to document: \`OnPremises\`, \`Online\`, or \`Both\` |
| \`OutputPath\` | String | No | Output directory for reports (default: current directory) |
| \`ExchangeServer\` | String | No* | FQDN of Exchange server (*required for OnPremises) |
| \`Credential\` | PSCredential | No | Credentials for authentication |
| \`TenantId\` | String | No | Azure AD Tenant ID for Exchange Online |
| \`AppId\` | String | No | Application ID for certificate-based auth |
| \`CertificateThumbprint\` | String | No | Certificate thumbprint for certificate-based auth |

## 📈 Report Outputs

### CSV Report
- Machine-readable format
- Each category as separate rows
- JSON-encoded data for complex objects
- Suitable for data analysis and automation

### HTML Report
- Professional, interactive interface
- Collapsible sections for easy navigation
- Environment-specific color coding
- Summary statistics dashboard
- Responsive design for various screen sizes
- Print-friendly formatting

### Report Categories

#### On-Premises Categories
- Organization Configuration
- Exchange Servers
- Mailbox Databases
- Database Availability Groups
- Receive/Send Connectors
- Transport Rules
- Virtual Directories
- Hybrid Configuration
- Client Access Services
- Mailbox Statistics

#### Exchange Online Categories
- Organization Configuration
- Tenant Information
- Mailbox Plans
- Transport Configuration
- Security Policies (EOP, Defender for Office 365)
- Mobile Device Policies
- Compliance Policies
- Detailed Statistics

## 🔧 Troubleshooting

### Common Issues

#### Module Installation Errors
\`\`\`powershell
# Run as Administrator if needed
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
\`\`\`

#### Connection Issues - Exchange Online
\`\`\`powershell
# Clear existing sessions
Get-PSSession | Remove-PSSession
Disconnect-ExchangeOnline -Confirm:\$false

# Reconnect
Connect-ExchangeOnline
\`\`\`

#### Connection Issues - On-Premises
\`\`\`powershell
# Verify WinRM configuration
winrm quickconfig

# Test connectivity
Test-NetConnection -ComputerName "exchange01.contoso.com" -Port 80
\`\`\`

#### Permission Issues
- Ensure proper administrative roles are assigned
- For Exchange Online: Global Admin or Exchange Admin
- For On-Premises: Organization Management role

### Error Handling
The script includes comprehensive error handling:
- Continues execution if individual data collection fails
- Logs warnings for failed operations
- Provides detailed error messages
- Ensures proper cleanup of connections

## 🔒 Security Considerations

### Authentication Methods
- **Interactive Authentication**: Prompts for credentials
- **Certificate-Based Authentication**: For unattended execution
- **Credential Objects**: For scripted scenarios

### Data Protection
- Reports may contain sensitive configuration data
- Store reports in secure locations
- Consider encryption for sensitive environments
- Review reports before sharing

### Network Security
- Uses encrypted connections (HTTPS/TLS)
- Supports certificate-based authentication
- No credentials stored in script

## 📅 Scheduling and Automation

### Task Scheduler Example
\`\`\`powershell
# Create scheduled task for monthly documentation
\$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\\Scripts\\Exchange-Documentation-Script.ps1 -Environment Online -OutputPath C:\\Reports"
\$trigger = New-ScheduledTaskTrigger -Monthly -At "02:00AM" -DaysOfMonth 1
\$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "Exchange Documentation" -Action \$action -Trigger \$trigger -Settings \$settings
\`\`\`

### Azure Automation Example
The script can be adapted for Azure Automation runbooks for cloud-based scheduling.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (\`git checkout -b feature/AmazingFeature\`)
3. Commit your changes (\`git commit -m 'Add some AmazingFeature'\`)
4. Push to the branch (\`git push origin feature/AmazingFeature\`)
5. Open a Pull Request

### Development Guidelines
- Follow PowerShell best practices
- Include error handling for new features
- Update documentation for new parameters
- Test with multiple Exchange versions

## 📝 Changelog

### 🔍 What's New in v3.0
- **Complete SMTP Relay Documentation**: Both on-premises and cloud connectors
- **Exchange Certificate Monitoring**: Expiration tracking and alerts
- **EWS Virtual Directory Coverage**: Complete client access documentation
- **Enhanced Security Reporting**: TLS, authentication, and certificate analysis
- **Critical Alert System**: Automatic identification of security issues
- **Comprehensive Transport Rules**: Complete mail flow rule documentation
- **Federation and Hybrid Details**: Organization relationships and sharing policies

### Version 2.0
- Enhanced Exchange Online support
- Added Microsoft Graph integration
- Improved HTML report design
- Added certificate-based authentication
- Enhanced error handling and logging
- Added module auto-installation

### Version 1.0
- Initial release
- Basic on-premises and Exchange Online support
- CSV and HTML report generation

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

### Getting Help
- Check the [Issues](https://github.com/yourusername/exchange-documentation-script/issues) page
- Review the troubleshooting section
- Ensure you have the latest version

### Reporting Issues
When reporting issues, please include:
- PowerShell version (\`\$PSVersionTable\`)
- Exchange environment details
- Error messages (full text)
- Steps to reproduce

### Feature Requests
Feature requests are welcome! Please:
- Check existing issues first
- Provide detailed use case
- Explain expected behavior

## 🙏 Acknowledgments

- Microsoft Exchange Team for comprehensive PowerShell cmdlets
- PowerShell community for best practices and examples
- Contributors and testers who helped improve the script

## 📞 Contact

- **Author**: [Your Name]
- **Email**: [your.email@domain.com]
- **GitHub**: [@yourusername](https://github.com/yourusername)
- **LinkedIn**: [Your LinkedIn Profile]

---

**⭐ If this script helps you, please consider giving it a star on GitHub!**
