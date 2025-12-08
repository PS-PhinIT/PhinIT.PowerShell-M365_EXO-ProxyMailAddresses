# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-08

### Added

#### Core Functionality
- PowerShell-based ProxyAddress manager with modern WPF GUI
- User search functionality (by DisplayName, SamAccountName, or Mail)
- Display all ProxyAddresses for selected user
- Add new ProxyAddresses with type selection (SMTP, smtp, SIP, X400, X500)
- Remove existing ProxyAddresses with confirmation
- Set primary SMTP address functionality
- Automatic demotion of existing primary when setting new primary
- Real-time status updates in status bar
- Refresh functionality to reload user data

#### User Interface
- Modern WPF interface with professional styling
- Blue theme with Microsoft-style design
- User list with search results
- ProxyAddresses data grid with sortable columns
- Type, Address, and Primary status display
- Expandable "Add New ProxyAddress" section
- Status bar with timestamp
- Confirmation dialogs for destructive operations
- Warning dialog when removing primary SMTP address
- Button states (enabled/disabled based on context)
- Keyboard shortcuts (Enter key support)

#### Validation & Error Handling
- Email format validation using .NET MailAddress class
- Duplicate ProxyAddress detection
- Empty input validation
- Proper error messages with context
- Try-catch blocks around all AD operations
- Graceful error recovery

#### Security Features
- Verification of Active Directory module availability
- Proper WPF threading model (Dispatcher.Invoke)
- No credential storage (uses current Windows credentials)
- Confirmation dialogs for all destructive operations
- Input validation and sanitization
- Limited AD query results (50 users max)
- Optimized AD queries with ResultSetSize parameter

#### Documentation
- Comprehensive README.md with feature overview
- QUICKSTART.md for rapid onboarding
- INSTALLATION.md with detailed setup instructions
- EXAMPLES.md with real-world usage scenarios
- CONTRIBUTING.md with contribution guidelines
- SECURITY.md with security considerations
- LICENSE file (MIT License)
- This CHANGELOG.md file

#### Configuration
- .gitignore for PowerShell projects
- Git-friendly file structure
- Professional repository setup

### Technical Details

#### PowerShell Features Used
- WPF/XAML for GUI
- ActiveDirectory module integration
- .NET Framework assemblies (PresentationFramework, PresentationCore, WindowsBase)
- Event-driven programming model
- Advanced parameter validation
- Proper error handling patterns

#### AD Operations
- Get-ADUser with filtered queries
- Set-ADUser for modifications
- Add, Remove, and Replace operations on ProxyAddresses
- Proper handling of multi-valued attributes

#### Code Quality
- Follows PowerShell best practices
- Clear function separation
- Comprehensive comments
- Consistent naming conventions
- Proper indentation (4 spaces)
- CmdletBinding support

### Fixed
- XAML inline event handlers removed (used Add_* methods instead)
- System.Windows.Forms.DoEvents() replaced with proper WPF Dispatcher
- AD query optimization (changed from Select-Object -First to -ResultSetSize)
- Email validation improved (regex replaced with .NET MailAddress)
- Removed unused System.Windows.Forms assembly reference

### Security
- Input validation prevents injection attacks
- Proper error handling prevents information disclosure
- WPF threading model prevents reentrancy issues
- Limited query results prevent resource exhaustion
- Confirmation dialogs prevent accidental data loss

## [Unreleased]

### Planned for v1.1.0
- Export/Import ProxyAddresses to/from CSV
- Batch operations for multiple users
- Advanced search filters
- Domain-specific filtering
- ProxyAddress templates
- Dark mode theme option

### Planned for v1.2.0
- Configuration file support
- Custom default values
- Logging to file
- Audit trail functionality
- Undo/Redo capability

### Planned for v2.0.0
- Multi-language support (English/German)
- MS Graph API integration (direct MS365 access)
- Advanced reporting features
- Scheduled tasks support
- PowerShell module packaging

### Under Consideration
- Digital code signing
- MSI installer
- Automatic updates
- Integration with ticketing systems
- Approval workflow
- Read-only mode
- Rollback functionality
- Pre-change backups

---

## Version History Overview

- **1.0.0** (2025-12-08): Initial release with core functionality

---

## How to Read This Changelog

- **Added**: New features
- **Changed**: Changes in existing functionality
- **Deprecated**: Soon-to-be removed features
- **Removed**: Removed features
- **Fixed**: Bug fixes
- **Security**: Security improvements

## Links

- [GitHub Repository](https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses)
- [Issues](https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses/issues)
- [Pull Requests](https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses/pulls)
