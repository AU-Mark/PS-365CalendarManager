# PS-365CalendarManager
<p align="center">
    <img src="https://raw.githubusercontent.com/AU-Mark/PS-365CalendarManager/refs/heads/main/Source%20Files/Calendar%20Manager.png" />
</p>

## Description
This script installs powershell files and configures a scheduled task that runs daily at 8am and processes all Active Directory users that are enabled without password never expires checked in their account options. Any account that has a password that is going to expire within the configurable number of days (Default is 14) will be sent an email to their email address. 

The email is dynamically crafted based on options configured during installation that are saved in a JSON config file. The JSON config file can be recreated if it's missing by deleting the config file and either running the installation script OR the main script in an interactive powershell session after installation.

## Dependencies/Prerequisites
### Dependencies
The following PowerShell modules are required and will be installed during installation to ensure the scripts run without error:
*   ExchangeOnlineManagement version 3.8.0

## Installation
### Run directly from github
```powershell
irm https://github.com/AU-Mark/PS-365CalendarManager/raw/refs/heads/main/Start-CalendarManager.ps1 | iex
```
### OR

### Download and execute the script
Download and execute Install-PasswordExpiryNotification.ps1. The main scripts are embedded within and installed during execution after the options are configured.
