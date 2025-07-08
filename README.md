# PS-365CalendarManager

<p align="center">
    <img src="https://raw.githubusercontent.com/AU-Mark/PS-365CalendarManager/refs/heads/main/Source%20Files/Calendar%20Manager.png" />
</p>

## Description

An interactive PowerShell script for managing Exchange Online calendar permissions with a modern, user-friendly interface. This tool provides comprehensive calendar permission management capabilities including viewing, adding, modifying, and removing calendar permissions for user, shared, and room mailboxes.

**Key Features:**
- **Interactive Menu System**: Navigate using arrow keys with Unicode-styled menus
- **Permission Management**: View, add, modify, and remove calendar permissions
- **Multiple Permission Levels**: Support for all Exchange Online calendar permission levels (Owner, Editor, Reviewer, etc.)
- **Delegate Support**: Configure delegate permissions with meeting invite forwarding
- **Email Notifications**: Send calendar sharing invitations with preview functionality
- **Mailbox Support**: Works with user mailboxes, shared mailboxes, and room mailboxes
- **Input Validation**: Comprehensive email format validation and mailbox existence checking
- **Visual Status Tracking**: Real-time status bar and window title updates
- **Error Handling**: Robust error handling with user-friendly messages

**Version:** 1.0  
**Release Date:** 2025-06-19  
**Author:** Mark Newton

## Dependencies/Prerequisites

### System Requirements
- **PowerShell**: Version 5.1 or later
- **Execution Policy**: Must allow script execution (Run as Administrator required)
- **Exchange Online**: Access to Exchange Online tenant with appropriate permissions

### Dependencies
The following PowerShell module is required and will be automatically installed if not present:
- **ExchangeOnlineManagement** version 3.8.0 or later

### Required Permissions
The executing user must have one of the following Exchange Online roles:
- Exchange Administrator
- Global Administrator
- Organization Management
- Or custom role with calendar permission management rights

## Usage

### Run directly from GitHub
```powershell
irm https://github.com/AU-Mark/PS-365CalendarManager/raw/refs/heads/main/Start-CalendarManager.ps1 | iex
```

### OR

### Download and execute the script
1. Download the `Start-CalendarManager.ps1` file
2. Open PowerShell as Administrator
3. Navigate to the script location
4. Execute: `.\Start-CalendarManager.ps1`

## Menu Options

### Main Menu
1. **📅 View Calendar Permissions** - Display current calendar permissions for a specified mailbox
2. **🆕 Add Calendar Permission** - Grant calendar access to a user with specified permission level
3. **✏️ Modify Calendar Permission** - Change existing calendar permissions for a user
4. **🗑️ Remove Calendar Permission** - Remove all calendar permissions for a user

### Permission Levels Supported
- **Owner**: Full control of the calendar
- **PublishingEditor**: Create, read, modify, and delete items and folders
- **Editor**: Create, read, modify, and delete items
- **PublishingAuthor**: Create and read items and folders; modify and delete own items
- **Author**: Create and read items; modify and delete own items
- **NonEditingAuthor**: Create and read items; delete own items
- **Reviewer**: Read items only
- **Contributor**: Create items only
- **AvailabilityOnly**: View free/busy information only
- **LimitedDetails**: View free/busy and subject information

### Delegate Features
When granting **Editor** permissions, additional delegate options are available:
- Standard Editor permissions only
- Delegate with meeting invite forwarding
- Delegate with private item access

## Navigation

- **Arrow Keys** (↑↓ or ←→): Navigate menu options
- **Enter**: Select highlighted option
- **Escape**: Cancel current operation
- **Back Options**: Available in sub-menus to return to previous screens

## Features in Detail

### Email Notifications
- Preview notification emails before sending
- Customizable notification settings per operation
- Professional Exchange Online-style calendar sharing invitations

### Input Validation
- Email format validation using regex and .NET MailAddress class
- Mailbox existence verification in Exchange Online
- Real-time feedback for invalid inputs

### Status Tracking
- Dynamic window title updates showing current operation
- Status bar displaying target mailbox, user, and current step
- Color-coded permission levels for easy identification

### Error Handling
- Graceful handling of connection issues
- Clear error messages with suggested resolutions
- Automatic cleanup and disconnection on exit

## Notes

- The script automatically handles Exchange Online connection and authentication
- Modern authentication (OAuth) is used for secure connections
- Sessions are automatically cleaned up on exit
- The script supports localized calendar folder names
- All operations include confirmation steps to prevent accidental changes

## Troubleshooting

**Module Installation Issues**: If the ExchangeOnlineManagement module fails to install, ensure you're running as Administrator and have internet connectivity.

**Connection Failures**: Verify your account has the required Exchange Online permissions and that modern authentication is enabled.

**Mailbox Not Found**: Ensure the mailbox exists and you have permission to view it. The script searches by email address, UPN, or display name.
