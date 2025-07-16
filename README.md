# PS-365CalendarManager

<div align="center">
  <img src="https://raw.githubusercontent.com/AU-Mark/PS-365CalendarManager/refs/heads/main/Source%20Files/Calendar%20Manager.png" alt="PS-365CalendarManager Logo" width="600"/>
</div>

<div align="center">

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![Exchange Online](https://img.shields.io/badge/Exchange_Online-0078D4?style=for-the-badge&logo=microsoft-exchange&logoColor=white)
![Version](https://img.shields.io/badge/Version-1.1-blue?style=for-the-badge)

</div>

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Requirements](#-requirements)
- [Installation & Usage](#-installation--usage)
- [Multi-Calendar Support](#-multi-calendar-support)
- [Permission Levels](#-permission-levels)
- [Navigation](#-navigation)
- [Feature Details](#-feature-details)
- [Troubleshooting](#-troubleshooting)
- [Version History](#-version-history)
- [Contributing](#-contributing)

## 🎯 Overview

An interactive PowerShell script for managing Exchange Online calendar permissions with a modern, user-friendly interface. This tool provides comprehensive calendar permission management capabilities including viewing, adding, modifying, and removing calendar permissions for user, shared, and room mailboxes.

**Perfect for:**
- Exchange Online administrators
- IT support teams managing calendar permissions
- Organizations requiring frequent calendar delegation changes
- Automating calendar permission workflows

## ✨ Features

### 🖥️ **Interactive Interface**
- **Arrow Key Navigation**: Modern menu system with ↑↓ or ←→ navigation
- **Unicode Styling**: Professional-looking menus with borders and colors
- **Real-time Status**: Dynamic window titles and status bars
- **Visual Feedback**: Color-coded permission levels and status indicators

### 📅 **Multi-Calendar Support** *(New in v1.1)*
- **Automatic Detection**: Discovers all calendars within a mailbox
- **Interactive Selection**: Choose from multiple calendars when available
- **Smart Auto-Selection**: Automatically proceeds with default calendar if only one exists
- **Subfolder Support**: Works with calendar subfolders and additional calendars

### 🔐 **Permission Management**
- **View Permissions**: Display current calendar permissions with color-coded levels
- **Add Permissions**: Grant calendar access with specified permission levels
- **Modify Permissions**: Update existing calendar permissions
- **Remove Permissions**: Remove all calendar access for users
- **Delegate Support**: Configure delegate permissions with meeting invite forwarding

### 📧 **Email Notifications**
- **Preview Functionality**: See sample notification emails before sending
- **Customizable Settings**: Choose whether to send notifications per operation
- **Professional Format**: Exchange Online-style calendar sharing invitations

### 🛡️ **Security & Validation**
- **Input Validation**: Comprehensive email format validation
- **Mailbox Verification**: Real-time mailbox existence checking
- **Error Handling**: Graceful handling of connection and permission issues
- **Secure Authentication**: Modern OAuth authentication with Exchange Online

## 📋 Requirements

### System Requirements
- **PowerShell**: Version 5.1 or later
- **Execution Policy**: Must allow script execution
- **Administrator Rights**: Required for module installation and Exchange Online connection

### Dependencies
The script will automatically install the following if not present:
- **ExchangeOnlineManagement** module (version 3.8.0 or later)

### Required Permissions
The executing user must have one of the following Exchange Online roles:
- Exchange Administrator
- Global Administrator
- Organization Management
- Custom role with calendar permission management rights

## 🚀 Installation & Usage

### Option 1: Direct Execution from GitHub
```powershell
# Run directly from GitHub (requires internet connection)
irm https://github.com/AU-Mark/PS-365CalendarManager/raw/refs/heads/main/Start-CalendarManager.ps1 | iex
```

### Option 2: Download and Execute
1. Download the `Start-CalendarManager.ps1` file
2. Open PowerShell as Administrator
3. Navigate to the script location
4. Execute the script:
```powershell
.\Start-CalendarManager.ps1
```

### Option 3: With Parameters
```powershell
# Specify admin UPN directly
.\Start-CalendarManager.ps1 -AdminUPN admin@contoso.com
```

## 📅 Multi-Calendar Support

The script now supports multiple calendars within a single mailbox:

### How It Works
1. **Detection**: Automatically scans for all calendar folders in the target mailbox
2. **Selection Menu**: Presents an interactive menu when multiple calendars are found
3. **Auto-Selection**: Proceeds automatically if only the default calendar exists
4. **Folder Support**: Works with calendar subfolders like "Personal", "Work", etc.

### Example Calendar Structure
```
├── Calendar (Default)
├── Personal Calendar
├── Work Projects
└── Team Meetings
```

### User Experience
- **Multiple Calendars**: Interactive selection menu appears
- **Single Calendar**: Automatically proceeds with default calendar
- **Visual Indicators**: Shows calendar names and item counts
- **Easy Navigation**: Arrow keys to select, Enter to confirm

## 🔑 Permission Levels

### Standard Permission Levels
| Permission Level | Description |
|------------------|-------------|
| **Owner** | Full control of the calendar |
| **PublishingEditor** | Create, read, modify, and delete items and folders |
| **Editor** | Create, read, modify, and delete items |
| **PublishingAuthor** | Create and read items and folders; modify and delete own items |
| **Author** | Create and read items; modify and delete own items |
| **NonEditingAuthor** | Create and read items; delete own items |
| **Reviewer** | Read items only |
| **Contributor** | Create items only |
| **AvailabilityOnly** | View free/busy information only |
| **LimitedDetails** | View free/busy and subject information |

### Delegate Features
When granting **Editor** permissions, additional delegate options are available:
- **Standard Editor**: Basic edit permissions
- **Delegate**: Editor permissions + meeting invite forwarding
- **Delegate with Private Access**: Full delegate permissions including private items

## 🎮 Navigation

### Menu Controls
- **Arrow Keys** (↑↓ or ←→): Navigate menu options
- **Enter**: Select highlighted option
- **Escape**: Cancel current operation
- **Back Options**: Available in sub-menus to return to previous screens

### Status Information
- **Window Title**: Shows current operation and context
- **Status Bar**: Displays target mailbox, user, and current step
- **Progress Indicators**: Visual feedback for long-running operations

## 🔧 Feature Details

### Email Notifications
- **Preview Mode**: See exactly what the notification email will look like
- **Customizable**: Choose to send or skip notifications per operation
- **Professional Format**: Uses Exchange Online's standard calendar sharing format
- **Multiple Languages**: Supports localized calendar folder names

### Input Validation
- **Email Format**: Regex and .NET MailAddress validation
- **Mailbox Existence**: Real-time verification in Exchange Online
- **Permission Validation**: Ensures valid permission combinations
- **Error Recovery**: Clear error messages with suggested fixes

### Error Handling
- **Connection Issues**: Graceful handling of Exchange Online connection problems
- **Permission Errors**: Clear explanations of permission-related issues
- **Module Dependencies**: Automatic detection and installation of required modules
- **Session Cleanup**: Automatic disconnection and cleanup on exit

## 🔧 Troubleshooting

### Common Issues

#### Module Installation Problems
**Issue**: ExchangeOnlineManagement module fails to install
**Solution**: 
- Ensure running as Administrator
- Check internet connectivity
- Try manual installation: `Install-Module -Name ExchangeOnlineManagement -Force`

#### Connection Failures
**Issue**: Cannot connect to Exchange Online
**Solution**:
- Verify account has required Exchange Online permissions
- Ensure modern authentication is enabled
- Check for conditional access policies blocking connection

#### Mailbox Not Found
**Issue**: "Mailbox not found" error
**Solution**:
- Verify mailbox exists and is accessible
- Check spelling of email address/UPN
- Ensure you have permission to view the target mailbox
- Try using the full UPN instead of just the email address

#### Permission Errors
**Issue**: Cannot modify calendar permissions
**Solution**:
- Verify you have Exchange Administrator role
- Check that the target user exists in the organization
- Ensure the calendar folder exists and is accessible

## 📈 Version History

### Version 1.1 (Current)
- ✅ **Multi-Calendar Support**: Automatic detection and selection of multiple calendars
- ✅ **Smart Auto-Selection**: Automatic calendar selection when only default exists

### Version 1.0
- ✅ Basic calendar permission management
- ✅ Exchange Online integration
- ✅ Support for all permission levels
- ✅ Delegate permission configuration
- ✅ Interactive menu system with arrow key navigation
- ✅ Email notification preview functionality
- ✅ Comprehensive input validation
- ✅ Real-time status tracking
- ✅ Enhanced error handling and user feedback

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📞 Support

If you encounter any issues or have questions:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review existing [Issues](https://github.com/AU-Mark/PS-365CalendarManager/issues)
3. Create a new issue with detailed information about your problem

---

<div align="center">
  
**Made with ❤️ for the Exchange Online community**

⭐ **Star this repository if you find it helpful!** ⭐

</div>
