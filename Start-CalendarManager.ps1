#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Interactive Exchange Online calendar permissions management tool with multi-calendar support.

.DESCRIPTION
    A comprehensive PowerShell script that provides an intuitive interface for managing Exchange Online 
    calendar permissions across user, shared, and room mailboxes. Features include:
    
    - Interactive menu-driven interface with arrow key navigation
    - Multi-calendar support with automatic detection and selection
    - Comprehensive permission management (view, add, modify, remove)
    - Input validation and error handling
    - Email notification options for permission changes
    - Support for delegate permissions and sharing flags
    - Real-time status tracking and progress indicators
    - Sample email previews for notification understanding
    
    The script automatically detects multiple calendars within a mailbox and allows users to select 
    which specific calendar to manage. If only the default calendar exists, it proceeds automatically.
    
    All operations are performed through Exchange Online PowerShell with proper authentication and 
    session management. The script includes comprehensive error handling and user-friendly feedback.

.PARAMETER AdminUPN
    Specifies the administrator User Principal Name (UPN) for connecting to Exchange Online.
    This account must have sufficient permissions to manage calendar permissions for the target mailboxes.
    
    If not provided, the script will prompt for the admin UPN during initialization.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None. This script provides interactive output to the console and does not return objects.

.EXAMPLE
    PS C:\> .\Start-CalendarManager.ps1 -AdminUPN admin@contoso.com
    
    Connects to Exchange Online using the specified admin account and launches the interactive 
    calendar permissions management interface.

.EXAMPLE
    PS C:\> .\Start-CalendarManager.ps1
    
    Launches the script and prompts for the admin UPN, then proceeds with the interactive interface.

.EXAMPLE
    PS C:\> Get-Help .\Start-CalendarManager.ps1 -Full
    
    Displays the complete help information for this script including all parameters and examples.

.NOTES
    Author:           Mark Newton
    Version:          1.1
    Created:          2025-06-15
    Last Modified:    2025-07-16
    
    Requirements:
    - PowerShell 5.1 or later
    - ExchangeOnlineManagement module (version 3.8.0 or later)
    - Exchange Online administrator permissions
    - Windows PowerShell execution policy allowing script execution
    
    Supported Mailbox Types:
    - User mailboxes
    - Shared mailboxes  
    - Room mailboxes
    - Equipment mailboxes
    
    Supported Permission Levels:
    - Owner, PublishingEditor, Editor, PublishingAuthor, Author
    - NonEditingAuthor, Reviewer, Contributor, AvailabilityOnly, LimitedDetails
    
    Features Added in v1.1:
    - Multi-calendar support and automatic detection
    - Interactive calendar selection menu
    
    Security Considerations:
    - Script requires administrative privileges
    - All operations are logged and tracked
    - Secure credential handling for Exchange Online connection
    - Input validation prevents common attack vectors

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/

.LINK
    https://docs.microsoft.com/en-us/exchange/recipients/mailbox-folder-permissions

.COMPONENT
    ExchangeOnlineManagement

.ROLE
    Exchange Administrator

.FUNCTIONALITY
    Exchange Online calendar permission management, mailbox administration, delegation management
#>

# ================================
# ===    CONFIGURATION VARS    ===
# ================================
#region Configuration Variables

# Script-level variables for consistent configuration
$Script:ModuleName = 'ExchangeOnlineManagement'
$Script:RequiredModuleVersion = '3.8.0'
$Script:IsConnected = $false

# Navigation and state management variables
$Script:MenuStack = [System.Collections.Generic.Stack[PSCustomObject]]::new()
$Script:CurrentAction = $null
$Script:CurrentTargetMailbox = $null
$Script:CurrentUserMailbox = $null
$Script:CurrentStep = $null

#endregion

# ================================
# ===        FUNCTIONS         ===
# ================================
#region Functions

Function Write-ColorEX {
    <#
    .SYNOPSIS
    Write-ColorEX is a wrapper around Write-Host delivering a lot of additional features for easier color options and logging.

    .DESCRIPTION
    Write-ColorEX is a wrapper around Write-Host delivering a lot of additional features for easier color options for native powershell, ANSI SGR, ANSI 4-bit color, and ANSI 8-bit color support. 
    It provides easy manipulation of colors, logging output to file (log) and nice formatting options out of the box.

    It provides:
    - Easy manipulation of colors
    - ANSI 4 color support with strings or integers
    - ANSI 8 color support with strings or integers
    - ANSI Text and Line Styles
    - Testing of ANSI support in your console if ANSI coloring or styles used
    - Logging output to file with optional parameters to log timestamps and log levels
    - Nice formatting options out of the box.
    - Ability to use aliases for a number of parameters

    .PARAMETER Text
    Text to display on screen and write to log file if specified.
    Accepts an array of strings.

    .PARAMETER Color
    Color of the text. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.

    Available native PWSH colors are: 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    DarkGray, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, DarkBlue

    Available ANSI4 colors are (if you supply a Dark color it will be converted into the Light color automatically): 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    LightGray, LightBlue, LightGreen, LightCyan, LightRed, LightMagenta, LightYellow, LightBlue, LightBlack

    More info on ANSI avilable at wikipedia
    https://en.wikipedia.org/wiki/ANSI_escape_code

    Available ANSI8 colors are: 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    LightGray, LightBlue, LightGreen, LightCyan, LightRed, LightMagenta, LightYellow, LightBlue, LightBlack
    DarkGray, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, DarkBlue

    More info on ANSI avilable at wikipedia
    https://en.wikipedia.org/wiki/ANSI_escape_code

    .PARAMETER BackGroundColor
    Color of the background. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.

    Available native PWSH colors are: 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    DarkGray, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, DarkBlue

    Available ANSI4 colors are (if you supply a Dark color it will be converted into the Light color automatically): 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    LightGray, LightBlue, LightGreen, LightCyan, LightRed, LightMagenta, LightYellow, LightBlue, LightBlack

    More info on ANSI avilable at wikipedia
    https://en.wikipedia.org/wiki/ANSI_escape_code

    Available ANSI8 colors are: 
    White, Green, Cyan, Red, Magenta, Yellow, Gray, Black, 
    LightGray, LightBlue, LightGreen, LightCyan, LightRed, LightMagenta, LightYellow, LightBlue, LightBlack
    DarkGray, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, DarkBlue

    More info on ANSI avilable at wikipedia
    https://en.wikipedia.org/wiki/ANSI_escape_code

    .PARAMETER ANSI4
    Switch to enable 4-bit ANSI color mode for terminals that support it. Enables the translation of color names to ANSI 4-bit color codes and the use of ANSI 4-bit color integers.

    .PARAMETER ANSI8
    Switch to enable 8-bit ANSI color mode for terminals that support it. Enables the translation of color names to ANSI 8-bit color codes and the use of ANSI 8-bit color integers.

    .PARAMETER Style
    Custom style parameters for ANSI-enabled terminals. Accepts an array of styles or an array of arrays of styles to apply to multiple text segments.

    .PARAMETER Bold
    Switch to make the whole line bold when using ANSI terminal support. Bold text in PowerShell is converted to to the lighter color value. 
    - For native PowerShell colors, that means you can only bold the Dark colored texts. Running bold on the regular colors will not show any difference. 
    - For ANSI 4-bit colors you can only bold the regular colors. Running bold on the light colors will not show any difference.
    Default is False.
    - For ANSI 8-bit colors, the regular and dark color names support bolding. Running bold on the light colors will not show any difference.

    .PARAMETER Faint
    Switch to make the whole line faint (decreased intensity) when using ANSI terminal support.

    .PARAMETER Italic
    Switch to make the whole line italic when using ANSI terminal support.

    .PARAMETER Underline
    Switch to underline the whole line when using ANSI terminal support.

    .PARAMETER Blink
    Switch to make the whole line blink when using ANSI terminal support.

    .PARAMETER CrossedOut
    Switch to display the whole line with a line through it (strikethrough) when using ANSI terminal support.

    .PARAMETER DoubleUnderline
    Switch to display the whole line with a double underline when using ANSI terminal support.

    .PARAMETER Overline
    Switch to display the whole line with a line above it when using ANSI terminal support.

    .PARAMETER StartTab
    Number of tabs to add before text. Default is 0.

    .PARAMETER LinesBefore
    Number of empty lines before text. Default is 0.

    .PARAMETER LinesAfter
    Number of empty lines after text. Default is 0.

    .PARAMETER StartSpaces
    Number of spaces to add before text. Default is 0.

    .PARAMETER LogFile
    Path to log file or name of the logfile. If only a filename is provided it will put in the LogPath directory; and extension of .log will be appended if no extension is provided.

    .PARAMETER LogPath
    Path to store the log file in if LogFile does not contain a path. If running in a script, it will default to PSScriptRoot. If running in console it will default to the current working directory.

    .PARAMETER DateTimeFormat
    Custom date and time format string. Default is yyyy-MM-dd HH:mm:ss

    .PARAMETER LogLevel
    The log level to include in the log file. Accepts a string. This is only provides options for writing to the log with log levels separate from the text. See logging example.

    .PARAMETER LogTime
    Switch to include the timestamp in the logfile

    .PARAMETER LogRetry
    Number of retries to write to log file, in case it can't write to it for some reason, before skipping. Default is 2.

    .PARAMETER Encoding
    Encoding of the log file. Default is Unicode.

    .PARAMETER ShowTime
    Switch to add time to console output. Default is not set.

    .PARAMETER NoNewLine
    Switch to not add new line at the end of the output. Default is not set.

    .PARAMETER HorizontalCenter
    Calculates the window width and inserts spaces to make the text center according to the present width of the powershell window. Default is false.

    .PARAMETER NoConsoleOutput
    Switch to not output to console. Default all output goes to console.

    .EXAMPLE
    # Writing text with multiple colors
    Write-ColorEX -Text 'Red ', 'Green ', 'Yellow ' -Color Red,Green,Yellow

    .EXAMPLE
    # Writing text with multiple colors and splitting text segments onto new lines for easier readability
    Write-ColorEX -Text 'This is text in Green ',
                      'followed by red ',
                      'and then we have Magenta... ',
                      "isn't it fun? "",
                      'Here goes DarkCyan' -Color Green,Red,Magenta,White,DarkCyan

    .EXAMPLE
    # Formatting with tabs, lines before and after
    Write-ColorEX -Text 'This could be a header with a blank line before and blank line after' -Color Cyan -LinesBefore 1 -LinesAfter 1
    Write-ColorEX -Text 'This is indented content' -Color White -StartTab 2
    Write-ColorEX -Text 'Back to normal indentation' -Color Gray -LinesAfter 1

    .EXAMPLE
    # Horizontal centering
    Write-ColorEX -Text 'This text could be a horiztonally centered header' -Color Green -HorizontalCenter -LinesBefore 1 -LinesAfter 1
    Write-ColorEX -Text 'Important ', 'Warning' -BackGroundColor DarkRed,DarkRed -HorizontalCenter -Bold

    .EXAMPLE
    # ANSI styling with different text effects
    Write-ColorEX -Text 'This text is bold' -Color DarkYellow -Bold
    Write-ColorEX -Text 'This text is italicized' -Color Green -Italic
    Write-ColorEX -Text 'This text is underlined' -Color Cyan -Underline
    Write-ColorEX -Text 'This text blinks' -Color Magenta -Blink
    Write-ColorEX -Text 'This text is crossed out' -Color Red -CrossedOut
    Write-ColorEX -Text 'This text has a double underline' -Color Blue -DoubleUnderline
    Write-ColorEX -Text 'This text has an overline' -Color White -Overline

    .EXAMPLE
    # Complex styling with different effects per text segment
    Write-ColorEX -Text "This segment is bold", " this one is italic", " this one blinks", " this one is crossed out" -Color Yellow,Cyan,Magenta,Red -Style Bold,Italic,Blink,CrossedOut

    .EXAMPLE
    # Applying multiple styles to different text segments using explicit array notation
    Write-ColorEX -Text 'This part is bold and italic', ' and this part is underlined and blinking' -Color DarkYellow,Cyan -Style @('Bold','Italic'),@('Underline','Blink')

    .EXAMPLE
    # ANSI4 color mode. This example shows how Red and DarkRed map to the same color.
    Write-ColorEX -Text 'ANSI4 Light Red ', 'ANSI4 Red ', 'ANSI4 Dark Red' -Color LightRed,Red,DarkRed -ANSI4
    Write-ColorEX -Text 'ANSI4 Light Red with Red Background ', 'ANSI4 Red with Light Red Background' -Color LightRed,Red -BackGroundColor DarkRed,LightRed -ANSI4

    # ANSI4 color mode with integers
    Write-ColorEX -Text 'ANSI4 Light Red ', 'ANSI4 Red ', 'ANSI4 Dark Red' -Color 91,31,31 -ANSI4
    Write-ColorEX -Text 'ANSI4 Light Red with Red Background ', 'ANSI4 Red with Light Red Background' -Color 91,31 -BackGroundColor 41,101 -ANSI4

    .EXAMPLE
    # ANSI8 color mode
    Write-ColorEX -Text 'ANSI 8 Light Red ', 'ANSI 8 Red ', 'ANSI 8 Dark Red' -Color LightRed,Red,DarkRed -ANSI8
    Write-ColorEX -Text 'ANSI 8 Light Red ', 'ANSI 8 Red ', 'ANSI 8 Dark Red' -Color LightRed,Red,DarkRed -BackGroundColor Red,DarkRed,LightRed -ANSI8

    # ANSI8 color mode with integers
    Write-ColorEX -Text 'ANSI 8 Light Red ', 'ANSI 8 Red ', 'ANSI 8 Dark Red' -Color 9,1,52 -ANSI8
    Write-ColorEX -Text 'ANSI 8 Light Red ', 'ANSI 8 Red ', 'ANSI 8 Dark Red' -Color 9,1,52 -BackGroundColor 1,52,9 -ANSI8

    .EXAMPLE
    # Creating menu options
    Write-ColorEX '1. ', 'View System Information   '-Color Yellow,Cyan -BackGroundColor Black -StartTab 2
    Write-ColorEX '2. ', 'Check Disk Space          ' -Color Yellow,Cyan -BackGroundColor Black -StartTab 2
    Write-ColorEX '3. ', 'Scan for Updates          ' -Color Yellow,Cyan -BackGroundColor Black -StartTab 2
    Write-ColorEX '4. ', 'Exit                      ' -Color Yellow,Cyan -BackGroundColor Black -StartTab 2

    .EXAMPLE
    # Writing color and reading input with a zero width space character so there arent two extra spaces after the text is outputted
    Write-ColorEX -Text "Enter the number of your choice: " -Color White -NoNewline -LinesBefore 1; $selected = Read-Host
    Write-ColorEX -Text "Are you sure you want to select $selected"," (Y/","N","): " -Color White,DarkYellow,Green,DarkYellow -NoNewline; $confirmed = Read-Host

    .EXAMPLE
    # Creating status indicators with different styles
    Write-ColorEX '[', 'SUCCESS', '] ' -Color White,Green,White -Style None,Bold,None
    Write-ColorEX '[', 'WARNING', '] ' -Color White,Yellow,White -Style None,Bold,None 
    Write-ColorEX '[', 'ERROR', '] ' -Color White,Red,White -Style None,Bold,None
    Write-ColorEX 'Operation completed with ', '1 ', 'success ', '2 ','warnings and ', '1 ', 'error' -Color White,Green,White,Yellow,White,Red,White

    .EXAMPLE
    # Creating native PWSH dotted line boxed content
    Write-ColorEX "+----------------------+" -Color Cyan
    Write-ColorEX "$([char]166)", " System Status Report ", "$([char]166)" -Color Cyan,White,Cyan
    Write-ColorEX "+----------------------+" -Color Cyan
    Write-ColorEX "$([char]166)", " CPU: ", "42%             ", "$([char]166)" -Color Cyan,White,Green,Cyan
    Write-ColorEX "$([char]166)", " Memory: ", "68%          ", "$([char]166)" -Color Cyan,White,Yellow,Cyan
    Write-ColorEX "$([char]166)", " Disk: ", "89%            ", "$([char]166)" -Color Cyan,White,Red,Cyan
    Write-ColorEX "+----------------------+" -Color Cyan

    .EXAMPLE
    # Creating ANSI solid line boxed content
    Write-ColorEX "|","                      ","|" -Color Cyan, Cyan, Cyan -HorizontalCenter -Style None,Overline,None
    Write-ColorEX "|", " System Status Report ", "|" -Color Cyan,White,Cyan -HorizontalCenter
    Write-ColorEX "|","                      ","|" -Color Cyan, Cyan, Cyan -HorizontalCenter -Style None,Underline,None
    Write-ColorEX "|", " CPU: ", "42%             ", "|" -Color Cyan,White,Green,Cyan -HorizontalCenter
    Write-ColorEX "|", " Memory: ", "68%          ", "|" -Color Cyan,White,Yellow,Cyan -HorizontalCenter
    Write-ColorEX "|", " Disk: ", "89%            ", "|" -Color Cyan,White,Red,Cyan -HorizontalCenter
    Write-ColorEX " ","                      "," " -Color Cyan, Cyan, Cyan -HorizontalCenter -Style None,Overline,None

    .EXAMPLE
    # Using with logging capabilities
    Write-ColorEX -Text "Initializing application..." -Color White  -ShowTime -LogFile "C:\Temp\Write-ColorEX.log" 
    Write-ColorEX -Text "Reading configuration..." -Color White  -ShowTime -LogFile "Write-ColorEX" 
    Write-ColorEX -Text "Configuration ", "loaded successfully" -Color White,Green  -ShowTime -LogFile "Write-ColorEX.log" -LogTime
    Write-ColorEX -Text "Running disk space check" -Color White -ShowTime -LogFile "Write-ColorEX.log" -LogPath "C:\Temp" -LogTime

    .EXAMPLE
    # Using with logging capabilities and log levels.
    # Note: If you use LogLevel and put the loglevel in the text it will show twice in the recorded log
    # This example uses LogLevel parameter and colors the whole line.
    Write-ColorEX -Text "Disk space running low" -Color Yellow  -ShowTime -LogFile "Write-ColorEX.log" -LogLevel "WARNING" -LogTime
    # This example includes the log level in the message instead of the parameter and colors the loglevel only.
    Write-ColorEX -Text "[WARNING] ","Disk space running low" -Color Yellow,Grey  -ShowTime -LogFile "Write-ColorEX.log" -LogTime     
    
    .EXAMPLE
    # Using parameter aliases
    Write-ColorEX -T "Starting ", "process" -C Gray,Blue -L "Write-ColorEX" -ShowTime
    Write-ColorEX -T "Process ", "completed" -C Gray,Green -L "Write-ColorEX.log" -ShowTime
    
    .EXAMPLE
    # Writing out to the log with specific text encoding
    Write-ColorEX -Text 'Testuję czy się ładnie zapisze, czy będą problemy' -Encoding unicode -LogFile 'C:\temp\testinggg.txt' -Color Red -NoConsoleOutput

    .NOTES
    Understanding Custom date and time format strings: https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
    Project support: https://github.com/EvotecIT/PSWriteColor
    Original idea: Josh (https://stackoverflow.com/users/81769/josh)

    #>
    [alias('Write-ColourEX')]
    [CmdletBinding()]
    param (
        [alias ('T')][string[]] $Text,
        [ValidateScript({$_ -is [string] -or $_ -is [int] -or $_ -is [int[]] -or $_ -is [string[]]})][alias ('C', 'ForegroundColor', 'FGC')][array] $Color = $null,
        [ValidateScript({$_ -is [string] -or $_ -is [int] -or $_ -is [int[]] -or $_ -is [string[]]})][alias ('B', 'BGC')][array] $BackGroundColor = $null,
        [alias ('A4')][switch] $ANSI4,
        [alias ('A8')][switch] $ANSI8,
        [ValidateScript({$_ -is [string] -or $_ -is [int] -or $_ -is [int[]] -or $_ -is [string[]] -or $_ -is [object[]]})][alias ('S')][object] $Style = $null,
        [switch] $Bold,
        [switch] $Faint,
        [switch] $Italic,
        [switch] $Underline,
        [switch] $Blink,
        [alias ('Strikethrough')][switch] $CrossedOut,
        [switch] $DoubleUnderline,
        [switch] $Overline,
        [alias ('Indent')][int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [int] $StartSpaces = 0,
        [alias ('L')][string] $LogFile = '',
        [alias ('LP')][string] $LogPath = $(If ($PSScriptRoot) {$PSScriptRoot} Else {$PWD.Path}),
        [alias ('LL', 'LogLvl')][string] $LogLevel = '',
        [alias ('LT')][switch] $LogTime,
        [Alias('DateFormat', 'TimeFormat', 'Timestamp', 'TS')][string] $DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
        [int] $LogRetry = 2,
        [ValidateSet('unknown', 'string', 'unicode', 'bigendianunicode', 'utf8', 'utf7', 'utf32', 'ascii', 'default', 'oem')][string]$Encoding = 'Unicode',
        [switch] $ShowTime,
        [switch] $NoNewLine,
        [alias('Center')][switch] $HorizontalCenter,
        [alias ('BL', 'Empty', 'Blank')][switch] $BlankLine,
        [alias('HideConsole', 'NoConsole', 'LogOnly', 'LO')][switch] $NoConsoleOutput
    )

    function Test-AnsiSupport {
        [CmdletBinding()]
        param()
    
        # Initialize collection for results
        $results = [PSCustomObject]@{
            IsAnsiSupported = $False
            Details = @{
                PowerShellVersion = $PSVersionTable.PSVersion.ToString()
                IsConsoleHost = $Host.Name -eq 'ConsoleHost' 
                HasVirtualTerminalProcessing = $False
                HasCompatibleTerminalEnv = $False
                IsPSCore = $PSVersionTable.PSVersion.Major -ge 6
                OperatingSystem = [System.Environment]::OSVersion.Platform
            }
        }
    
        # VT Processing is automatically enabled in PowerShell 7+
        If ($results.Details.IsPSCore) {
            $results.Details.HasVirtualTerminalProcessing = $true
        }
        # Check for console host (ISE doesn't support ANSI)
        elseIf (-not $results.Details.IsConsoleHost) {
            $results.Details.HasVirtualTerminalProcessing = $False
        }
        # Use P/Invoke for Windows PowerShell to check console mode
        elseIf ($results.Details.OperatingSystem -eq 'Win32NT') {
            If ([System.Environment]::OSVersion.Version.Major -eq 10 -and [System.Environment]::OSVersion.Version.Build -ge 16257) {
                try {
                    # Define P/Invoke signatures
                    If (-not ('ConsoleHelper' -as [type])) {
                        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ConsoleHelper {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);}
"@ -ErrorAction SilentlyContinue
                    }
        
                    If ('ConsoleHelper' -as [type]) {
                        # Constants
                        $STDOUT_HANDLE = -11
                        $ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
                        
                        # Get console mode
                        $stdoutHandle = [ConsoleHelper]::GetStdHandle($STDOUT_HANDLE)
                        $consoleMode = 0
                        
                        If ([ConsoleHelper]::GetConsoleMode($stdoutHandle, [ref]$consoleMode)) {
                            $results.Details.HasVirtualTerminalProcessing = ($consoleMode -band $ENABLE_VIRTUAL_TERMINAL_PROCESSING) -ne 0
                        }
                    }
                }
                Catch {
                    # P/Invoke failed, continue with other checks
                    $results.Details.HasVirtualTerminalProcessing = $False
                }
            } Else {
                Write-Warning 'PowerShell is not capable of ANSI support on versions if Windows 10 earlier than build 16257'
                Return $False
            }
        }
    
        # Check environment variables that might indicate ANSI support
        $termEnv = [Environment]::GetEnvironmentVariable('TERM')
        $colorTerm = [Environment]::GetEnvironmentVariable('COLORTERM')
        $conEmuANSI = [Environment]::GetEnvironmentVariable('ConEmuANSI')
        
        # Linux/macOS terminals or Windows terminals like ConEmu/cmder
        $results.Details.HasCompatibleTerminalEnv = (
            -not [string]::IsNullOrEmpty($termEnv) -or
            -not [string]::IsNullOrEmpty($colorTerm) -or
            -not [string]::IsNullOrEmpty($conEmuANSI)
        )
    
        # Make the first determination based on system checks
        $results.IsAnsiSupported = $results.Details.HasVirtualTerminalProcessing -or $results.Details.HasCompatibleTerminalEnv
  
        If (-not ($results.IsAnsiSupported)) {
            If (-not $ansiSupport.Details.IsConsoleHost) {
                Write-Warning 'Reason: Not running in compatible console host'
            } elseIf (-not ($ansiSupport.Details.HasVirtualTerminalProcessing -or $ansiSupport.Details.HasCompatibleTerminalEnv)) {
                If ($results.Details.OperatingSystem -eq 'Win32NT' -and -not ($ansiSupport.Details.HasVirtualTerminalProcessing)) {
                    If ([System.Environment]::OSVersion.Version.Major -eq 10 -and [System.Environment]::OSVersion.Version.Build -ge 16257) {
                        # Enable VirtualTerminalProcessing for the current session only
                        # Define P/Invoke signatures
                        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ConsoleHelper {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
}
"@

                        # Constants
                        $STDOUT_HANDLE = -11
                        $ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
                        
                        # Get console mode
                        $stdoutHandle = [ConsoleHelper]::GetStdHandle($STDOUT_HANDLE)
                        $consoleMode = 0
                        
                        If ([ConsoleHelper]::GetConsoleMode($stdoutHandle, [ref]$consoleMode)) {
                            # Enable VirtualTerminalProcessing
                            $consoleMode = $consoleMode -bor $ENABLE_VIRTUAL_TERMINAL_PROCESSING
                            $result = [ConsoleHelper]::SetConsoleMode($stdoutHandle, $consoleMode)
                            

                            If (-not $result) {
                                Write-Warning 'ANSI is not supported: Virtual terminal processing is not enabled and we were unable to enable it'
                                Return $False
                            } Else {
                                Write-Warning 'ANSI is supported and has been enabled for this console session only. Recommend enabling it permanently.'
                                Return $True
                            }
                        }
                        Write-Warning 'ANSI support could not be automatically enabled for this console session. Recommend enabling it permanently.'
                        Return $False
                    }
                }
                Write-Warning 'ANSI is not supported: Virtual terminal processing is not enabled'
                Return $False
            }
        } Else {
            Return $True
        }
    }

    $ANSISupport = Test-AnsiSupport
    If (-not $ANSISupport) {
        $Style = @()
        $ANSI4 = $False
        $ANSI8 = $False
    }

    # If we are writing out to the console, skip all console related sections
    If (-not $NoConsoleOutput) {
        # ESC sequences to initiate ANSI styling
        $esc=$([char]27)

        # Hashtable of ANSI styles
        $ANSI = @{
            'Reset' = "$esc[0m"
            'Bold' = "$esc[1m"
            'Faint' = "$esc[2m"
            'Italic' = "$esc[3m"
            'Underline' = "$esc[4m"
            'Blink' = "$esc[5m"
            'CrossedOut' = "$esc[9m"
            'DoubleUnderline' = "$esc[21m"
            'Overline' = "$esc[53m"
            'None' = ""
        }

        $Colors = @{
            # Native Color Families
            # Neutral family
            Black = @('Black',30,40,0);
            LightBlack = @('DarkGray',90,100,238); 
            DarkGray = @('DarkGray',90,100,8); 
            Gray = @('Gray',37,107,7);
            LightGray = @('Gray',37,47,253); 
            White = @('White',97,107,15);
            
            # Red family
            DarkRed = @('DarkRed',31,41,52); 
            Red = @('Red',31,41,1); 
            LightRed = @('Red',91,101,9); 
            
            # Green family
            DarkGreen = @('DarkGreen',32,42,28); 
            Green = @('Green',32,42,2); 
            LightGreen = @('Green',92,102,10);
            
            # Yellow family
            DarkYellow = @('DarkYellow',33,43,136); 
            Yellow = @('Yellow',33,43,220); 
            LightYellow = @('Yellow',93,103,11);
            
            # Blue family
            DarkBlue = @('DarkBlue',34,44,19); 
            Blue = @('Blue',34,44,4); 
            LightBlue = @('Blue',94,104,12);
            
            # Magenta family
            DarkMagenta = @('DarkMagenta',35,45,53); 
            Magenta = @('Magenta',35,45,5); 
            LightMagenta = @('Magenta',95,105,13);
            
            # Cyan family
            DarkCyan = @('DarkCyan',36,46,30); 
            Cyan = @('Cyan',36,46,6); 
            LightCyan = @('Cyan',96,106,14);
            
            # Additional ANSI 8-bit Color Families
            # Orange family
            DarkOrange = @('DarkYellow',33,43,166); 
            Orange = @('DarkYellow',33,43,208); 
            LightOrange = @('Yellow',33,43,215);
            
            # Purple family
            DarkPurple = @('DarkMagenta',35,45,54); 
            Purple = @('DarkMagenta',35,45,93); 
            LightPurple = @('Magenta',35,45,135);
            
            # Pink family
            DarkPink = @('DarkMagenta',35,45,163); 
            Pink = @('Magenta',35,45,205); 
            LightPink = @('Magenta',95,105,218);
            
            # Brown family
            DarkBrown = @('DarkRed',31,41,88); 
            Brown = @('DarkRed',31,41,130); 
            LightBrown = @('DarkYellow',33,43,173);
            
            # Teal family
            DarkTeal = @('DarkCyan',36,46,23); 
            Teal = @('DarkCyan',36,46,30); 
            LightTeal = @('Cyan',36,46,80);
            
            # Violet family
            DarkViolet = @('DarkMagenta',35,45,128); 
            Violet = @('Magenta',35,45,134); 
            LightViolet = @('Magenta',95,105,177);
            
            # Lime family
            DarkLime = @('DarkGreen',32,42,34); 
            Lime = @('Green',32,42,118); 
            LightLime = @('Green',92,102,119);
            
            # Slate family
            DarkSlate = @('DarkGray',90,100,238); 
            Slate = @('Gray',37,107,102); 
            LightSlate = @('Gray',37,107,103);
            
            # Gold family
            DarkGold = @('DarkYellow',33,43,136); 
            Gold = @('Yellow',33,43,178); 
            LightGold = @('Yellow',93,103,185);
            
            # Sky family
            DarkSky = @('DarkBlue',34,44,24); 
            Sky = @('Blue',34,44,111); 
            LightSky = @('Cyan',36,46,152);
            
            # Coral family
            DarkCoral = @('DarkRed',31,41,167); 
            Coral = @('Red',31,41,209); 
            LightCoral = @('Red',91,101,210);
            
            # Olive family
            DarkOlive = @('DarkGreen',32,42,58); 
            Olive = @('DarkYellow',33,43,100); 
            LightOlive = @('DarkYellow',33,43,107);
            
            # Lavender family
            DarkLavender = @('DarkMagenta',35,45,97); 
            Lavender = @('Magenta',35,45,183); 
            LightLavender = @('Magenta',95,105,189);
            
            # Mint family
            DarkMint = @('DarkGreen',32,42,29); 
            Mint = @('Green',32,42,121); 
            LightMint = @('Green',92,102,157);
            
            # Salmon family
            DarkSalmon = @('DarkRed',31,41,173); 
            Salmon = @('Red',31,41,174); 
            LightSalmon = @('Red',91,101,175);
            
            # Indigo family
            DarkIndigo = @('DarkBlue',34,44,17); 
            Indigo = @('DarkMagenta',35,45,54); 
            LightIndigo = @('Blue',34,44,61);
            
            # Turquoise family
            DarkTurquoise = @('DarkCyan',36,46,31); 
            Turquoise = @('Cyan',36,46,43); 
            LightTurquoise = @('Cyan',96,106,86);
            
            # Ruby family
            DarkRuby = @('DarkRed',31,41,52); 
            Ruby = @('Red',31,41,124); 
            LightRuby = @('Red',91,101,161);
            
            # Jade family
            DarkJade = @('DarkGreen',32,42,22); 
            Jade = @('DarkGreen',32,42,35); 
            LightJade = @('Green',32,42,79);
            
            # Amber family
            DarkAmber = @('DarkYellow',33,43,130); 
            Amber = @('Yellow',33,43,214); 
            LightAmber = @('Yellow',93,103,221);
            
            # Steel family
            DarkSteel = @('DarkGray',90,100,60); 
            Steel = @('Gray',37,107,66); 
            LightSteel = @('White',97,47,146);
            
            # Crimson family
            DarkCrimson = @('DarkRed',31,41,88); 
            Crimson = @('Red',31,41,160); 
            LightCrimson = @('Red',91,101,161);
            
            # Emerald family
            DarkEmerald = @('DarkGreen',32,42,22); 
            Emerald = @('Green',32,42,36); 
            LightEmerald = @('Green',92,102,85);
            
            # Sapphire family
            DarkSapphire = @('DarkBlue',34,44,18); 
            Sapphire = @('Blue',34,44,25); 
            LightSapphire = @('Blue',94,104,69);
            
            # Wine family
            DarkWine = @('DarkRed',31,41,52); 
            Wine = @('DarkRed',31,41,88); 
            LightWine = @('Red',31,41,125);
        }

        If ($BlankLine -and ($Text.Length -eq 0)) {
            # Override these parameters if they were provided
            $HorizontalCenter = $False
            $StartTab = 0
            $StartSpaces = 0
            $ShowTime = $False
            # Add SPACES to fill the whole console
            $WindowWidth = $Host.UI.RawUI.BufferSize.Width
            $Text = [string[]]@()
            $Text += ' '
            For ($i = 1; $i -lt $WindowWidth; $i++) {
                $Text[0] += ' '
            } 
        } ElseIf ($BlankLine -and $Text.Length -gt 0) {
            # Override these parameters if they were provided
            $HorizontalCenter = $False
            $StartTab = 0
            $StartSpaces = 0
            $ShowTime = $False
            # Override the text and add SPACES to fill the whole console
            $WindowWidth = $Host.UI.RawUI.BufferSize.Width
            $Text = [string[]]@()
            $Text += ' '
            For ($i = 1; $i -lt $WindowWidth; $i++) {
                $Text[0] += ' '
            } 
        }
    
        # If Colors were provided, let's validate the colors over assign a default color
        If ($Color) {
            For ($i = 0; $i -lt $Text.Length; $i++) {
                If ($Color[$i] -is [string]) {
                    If ($Colors.Keys -NotContains $Color[$i]) {
                        If ($Color[$i] -ne 'None') {
                            Write-Warning "$($Color[$i]) is not a supported color string. The color gray will be assigned instead."
                        }
                        # Assign the default color of gray
                        $Color[$i] = 'Gray'
                    }
                    # Assign the default color equal to the first color provided
                    $DefaultColor = $Color[0]
                } ElseIf ($Color[$i] -is [Int32]) {
                    If ($ANSI8) {
                        If ($Color[$i] -notin 0..255) {
                            Write-Warning 'ANSI8 color integers must be between 0 and 255. The color gray will be assigned instead.'
                            # Assign the default color of gray
                            $Color[$i] = 7
                        }
                        $DefaultColor = $Color[0]
                    } ElseIf ($ANSI4) {
                        If ($Color[$i] -notin 30..37 -and $Color[$i] -notin 90..97) {
                            Write-Warning 'ANSI4 color integers must be between 30 and 37 or between 90 and 97. The color white will be assigned instead.'
                            # Assign the default color of gray
                            $Color[$i] = 37
                        }
                        # Assign the default color equal to the first color provided
                        $DefaultColor = $Color[0]
                    } Else {
                        Write-Error 'Integers can only be used for colors if using ANSI coloring. The color grey will be assigned instead.'
                        $Color[$i] = 'Gray'
                        return
                    }
                } ElseIf ($null -ne $Color[$i]) {
                    Write-Error 'Color must be a string or if using ANSI coloring an integer. Terminated.'
                    return
                } Else {
                    # Assign the default color of gray
                    $DefaultColor = $Color[0]
                }
            }
        } Else {
            # Assign default color of gray
            $DefaultColor = 'Gray'
        }

        # If background colors were provided, let's validate the colors and default to no background
        If ($BackGroundColor) {
            For ($i = 0; $i -lt $Text.Length; $i++) {
                If ($BackGroundColor[$i] -is [string]) {
                    If ($BackGroundColor[$i] -eq "None") {
                        # Default to no background color
                        $BackGroundColor[$i] = $Null
                    } ElseIf ($Colors.Keys -NotContains $BackGroundColor[$i]) {
                        Write-Warning "$($BackGroundColor[$i]) is not a supported background color string. No background color will being applied."
                        # Default to no background color
                        $BackGroundColor[$i] = $Null
                    }
                } ElseIf ($BackGroundColor[$i] -is [Int32]) {
                    If ($ANSI8) {
                        If ($BackGroundColor[$i] -notin 0..255) {
                            Write-Warning 'ANSI8 background color integers must be between 0 and 255. No background color will being applied.'
                            # Default to no background color
                            $BackGroundColor[$i] = $Null
                        }
                    } ElseIf ($ANSI4) {
                        If ($BackGroundColor[$i] -notin 40..47 -and $BackGroundColor[$i] -notin 100..107) {
                            Write-Warning 'ANSI4 background color integers must be between 30 and 37 or between 90 and 97. No background color will being applied.'
                            # Default to no background color
                            $BackGroundColor[$i] = $Null
                        }
                    } Else {
                        Write-Error 'Integers can only be used for background colors if using ANSI coloring. No background color will being applied.'
                        $BackGroundColor[$i] = $Null
                        return
                    }
                } ElseIf ($null -ne $BackGroundColor[$i]) {
                    Write-Error 'Color must be a string or if using ANSI coloring an integer. Terminated.'
                    return
                } 
            }
        }

        # If the line is bolded and using strings with ANSI8 coloring, automatically "bold" the line
        If ($ANSI8 -and $Bold) {
            For ($c = 0; $c -lt $Color.Length; $c++) {
                If ($Color[$c] -is [string]) {
                    If ($Color[$c] -like 'Dark*') {
                        $Color[$c] = $Color[$c].Substring(4)
                    } ElseIf ($Color[$c] -notlike 'Dark*' -and $Color[$c] -notlike 'Light*') {
                        $Color[$c] = "Light$($Color[$c])"
                    }
                }
            }
        }

        # Validate text styles if they were applied to individual text segments
        If ($Style) {
            # Store the original invocation line for analysis
            $InvocationLine = $MyInvocation.Line
            
            # Determine if the Style parameter was passed using explicit @() array syntax
            $ExplicitArrayNotation = $InvocationLine -match '-Style\s+@\('
            $ExplicitVariableNotation = $InvocationLine -match '-Style\s+\$[a-zA-Z0-9_]+'
            
            # Check if we're dealing with a style array or nested array
            $StyleIsArrayOfArrays = $false
            
            If ($Style.Length -gt 1) {
                If ($Style[0] -is [array] -or ($Style[1] -is [array] -and $Style[1].Length -gt 0)) {
                    $StyleIsArrayOfArrays = $true
                }
            }
            
            # Determine style application approach based on notation and content
            $ApplyMultipleStylesToFirstSegment = $false
            If ($ExplicitArrayNotation -or $ExplicitVariableNotation) {
                # Style was passed as @(...) or $variable, and is not an array of arrays
                If (-not $StyleIsArrayOfArrays -and $Style.Length -gt 1 -and $Text.Length -gt 1) {
                    $ApplyMultipleStylesToFirstSegment = $true
                }
            }
            
            # Structure the Style array based on how we determined it should be applied
            If ($ApplyMultipleStylesToFirstSegment) {
                # The user intended all styles to apply to the first text segment
                $ConvertedStyle = @()
                $ConvertedStyle += ,$Style  # Add all styles as an array for the first segment
                
                # Add empty style for remaining text segments
                For ($i = 1; $i -lt $Text.Length; $i++) {
                    $ConvertedStyle += @()
                }
                
                $Style = $ConvertedStyle
                $StyleIsArrayOfArrays = $true
            }
            ElseIf (-not $StyleIsArrayOfArrays -and $Style.Length -gt 0) {
                # Handle single string or simple array of styles by converting to array of arrays
                If ($Style -is [string]) {
                    # Single style string - apply to first text segment only
                    $ConvertedStyle = @()
                    $ConvertedStyle += ,@($Style)  # Add as single-element array for first segment
                    
                    # Add empty style for remaining text segments
                    For ($i = 1; $i -lt $Text.Length; $i++) {
                        $ConvertedStyle += @()
                    }
                    
                    $Style = $ConvertedStyle
                    $StyleIsArrayOfArrays = $true
                }
                ElseIf ($Style -is [array] -and $Text.Length -gt 1 -and $Style.Length -le $Text.Length) {
                    # Array of styles with one per text segment (one-to-one mapping)
                    $ConvertedStyle = @()
                    
                    For ($i = 0; $i -lt $Style.Length; $i++) {
                        If ($Style[$i] -is [array]) {
                            $ConvertedStyle += ,$Style[$i]
                        }
                        Else {
                            $ConvertedStyle += ,@($Style[$i])
                        }
                    }
                    
                    # Add empty styles for any remaining segments
                    For ($i = $Style.Length; $i -lt $Text.Length; $i++) {
                        $ConvertedStyle += @()
                    }
                    
                    $Style = $ConvertedStyle
                    $StyleIsArrayOfArrays = $true
                }
            }
            
            # Validate each style against supported styles
            For ($i = 0; $i -lt $Style.Length; $i++) {
                If ($Style[$i] -is [array]) {
                    For ($j = 0; $j -lt $Style[$i].Length; $j++) {
                        If ($ANSI.Keys -notcontains $Style[$i][$j]) {
                            Write-Warning "$($Style[$i][$j]) is not a supported ANSI style. It will be overwritten to cancel ANSI styling."
                            $Style[$i][$j] = $ANSI['None']
                        }
                        
                        # Handle ANSI8 bold styling for individual text segments
                        If ($ANSI8 -and $Style[$i][$j] -eq 'Bold' -and (-not $Bold)) {
                            For ($c = 0; $c -lt $Color.Length; $c++) {
                                If ($c -eq $i -and $Color[$c] -is [string]) {
                                    If ($Color[$c] -like 'Dark*') {
                                        $Color[$c] = $Color[$c].Substring(4)
                                    }
                                    ElseIf ($Color[$c] -notlike 'Dark*' -and $Color[$c] -notlike 'Light*') {
                                        $Color[$c] = "Light$($Color[$c])"
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    
        # Check if we're using ANSI4 or ANSI8 coloring, if not skip for efficiency
        If ($ANSI4 -or $ANSI8) {
            # Process through any color strings and convert them
            If ($Color) {
                For ($i = 0; $i -lt $Text.Length; $i++) {
                    If ($ANSI8) {
                        # If its a string, get the ANSI8 mapping
                        If ($Color[$i] -is [string]) {
                            # If the string is in the hashtable, assign the ANSI8 color integer
                            If ($Colors.Keys -Contains $Color[$i]) {
                                $Color[$i] = $Colors[$Color[$i]][3]
                            # Else assign a default color of White
                            } Else {
                                $Color[$i] = 15
                            }
                        } 
                    } ElseIf ($ANSI4) {
                        # If its a string, get the ANSI4 mapping
                        If ($Color[$i] -is [string]) {
                            # If the string is in the hashtable, assign the ANSI4 color integer
                            If ($Colors.Keys -Contains $Color[$i]) {
                                $Color[$i] = $Colors[$Color[$i]][1]
                            # Else assign a default color of White
                            } Else {
                                $Color[$i] = 37
                            }
                        } 
                    } 
                }
            }
    
            # Process through any background color strings and convert them
            If ($BackGroundColor) {
                For ($i = 0; $i -lt $Text.Length; $i++) {
                    If ($ANSI8) {
                        # If its a string, get the ANSI8 mapping
                        If ($BackGroundColor[$i] -is [string]) {
                            # If the string is in the hashtable, assign the ANSI8 color integer
                            If ($Colors.Keys -Contains $BackGroundColor[$i]) {
                                $BackGroundColor[$i] = $Colors[$BackGroundColor[$i]][3]
                            # Else not default background was applied, add a null value to write without a background color
                            } Else {
                                $BackGroundColor[$i] = $null
                            }
                        }
                    } ElseIf ($ANSI4) {
                        # If its a string, get the ANSI4BG mapping
                        If ($BackGroundColor[$i] -is [string]) {
                            # If the string is in the hashtable, assign the ANSI4BG color integer
                            If ($Colors.Keys -Contains $BackGroundColor[$i]) {
                                $BackGroundColor[$i] = $Colors[$BackGroundColor[$i]][2]
                            # Else not default background was applied, add a null value to write without a background color
                            } Else {
                                $BackGroundColor[$i] = $null
                            }
                        }
                    } 
                }
            }
        }

        If ($LinesBefore -ne 0) {
            For ($i = 0; $i -lt $LinesBefore; $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        } # Add empty line before
        If ($HorizontalCenter) {
            $MessageLength = 0
            ForEach ($Value in $Text) {
                $MessageLength += $Value.Length
            }
        
            $WindowWidth = $Host.UI.RawUI.BufferSize.Width
            $CenterPosition = [Math]::Max(0, $WindowWidth / 2 - [Math]::Floor($MessageLength / 2))
        
            # Only write spaces to the console if window width is greater than the message length
            If ($WindowWidth -ge $MessageLength) {
                Write-Host ("{0}" -f (' ' * $CenterPosition)) -NoNewline
            }
        } # Center the line horizontally according to the powershell window size. Ignore if blank line.
        If ($StartTab -ne 0) {
            For ($i = 0; $i -lt $StartTab; $i++) {
                Write-Host -Object "`t" -NoNewline 
            } 
        }  # Add TABS before text
        If ($StartSpaces -ne 0) {
            For ($i = 0; $i -lt $StartSpaces; $i++) {
                Write-Host -Object ' ' -NoNewline 
            } 
        }  # Add SPACES before text. Ignore if writing a blank line. 
        If ($ShowTime) {
            Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline -ForegroundColor DarkGray
        } # Add Time before output
        If ($Text.Count -ne 0) {
            If ($Color.Count -ge $Text.Count) {
                # the real deal coloring
                If ($null -eq $BackGroundColor) {
                    For ($i = 0; $i -lt $Text.Length; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($Color[$i])m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($Color[$i])m"
                        } Else {
                            $Parameters['ForegroundColor'] = $Color[$i]
                        }

                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                } Else {
                    For ($i = 0; $i -lt $Text.Length; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($Color[$i])m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($Color[$i])m"
                        } Else {
                            $Parameters['ForegroundColor'] = $Color[$i]
                        }

                        # Set the background color based on the type of coloring for the line
                        If ($null -ne $BackGroundColor[$i]) {
                            If ($ANSI8) {
                                $Parameters['Object'] += "$esc[48;5;$($BackGroundColor[$i])m"
                            } ElseIf ($ANSI4) {
                                $Parameters['Object'] += "$esc[$($BackGroundColor[$i])m"
                            } Else {
                                $Parameters['BackgroundColor'] = $BackGroundColor[$i]
                            }
                        }
                        
                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                }
            } Else {
                If ($null -eq $BackGroundColor) {
                    For ($i = 0; $i -lt $Color.Length ; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($Color[$i])m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($Color[$i])m"
                        } Else {
                            $Parameters['ForegroundColor'] = $Color[$i]
                        }

                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                    For ($i = $Color.Length; $i -lt $Text.Length; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color to the default color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($DefaultColor)m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($DefaultColor)m"
                        } Else {
                            $Parameters['ForegroundColor'] = $DefaultColor
                        }

                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                }
                Else {
                    For ($i = 0; $i -lt $Color.Length ; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($Color[$i])m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($Color[$i])m"
                        } Else {
                            $Parameters['ForegroundColor'] = $Color[$i]
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($null -ne $BackGroundColor[$i]) {
                            If ($ANSI8) {
                                $Parameters['Object'] += "$esc[48;5;$($BackGroundColor[$i])m"
                            } ElseIf ($ANSI4) {
                                $Parameters['Object'] += "$esc[$($BackGroundColor[$i])m"
                            } Else {
                                $Parameters['BackgroundColor'] = $BackGroundColor[$i]
                            }
                        }

                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                    For ($i = $Color.Length; $i -lt $Text.Length; $i++) {
                        # Initiate parameters for Write-Host
                        $Parameters = @{
                            'Object' = ''
                            'NoNewLine' = $True
                        }

                        If ($ANSISupport) {
                            # If individual text styles were applied, loop through them and apply them to the string
                            If ($Style) {
                                # We have detected that there are more than 1 text segment but only a single array of styles was provided. Apply them only to the first text segment.
                                If ($Text.Length -gt 1 -and $Style.Length -gt 1 -and $Text.Length -ne $Style.Length -and $Style[0] -isnot [array] -and $Style[0] -isnot [char]) {
                                    If ($i -eq 1) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    }
                                } Else {
                                    # For a single string with an array of text styles, we have to loop through each style to apply them
                                    If ($Text.Length -eq 1 -and $Style -is [array]) {
                                        For ($s = 0; $s -lt $Style.Length; $s++) {
                                            $Parameters['Object'] += $ANSI[$Style[$s]]
                                        }
                                    # Else loop through the styles based on if its an array or a string
                                    } Else {
                                        If ($Style[$i] -is [array]) {
                                            ForEach ($TextStyle in $Style[$i]) {
                                                $Parameters['Object'] += $ANSI[$TextStyle]
                                            }
                                        } ElseIf ($Style[$i] -is [string]) {
                                            $Parameters['Object'] += $ANSI[$Style[$i]]
                                        }
                                    }
                                }
                            }

                            # If line styles were applied, apply each one to the whole line
                            If ($Bold) { $Parameters['Object'] += $ANSI['Bold'] }
                            If ($Faint) { $Parameters['Object'] += $ANSI['Faint'] }
                            If ($Italic) { $Parameters['Object'] += $ANSI['Italic'] }
                            If ($Underline) { $Parameters['Object'] += $ANSI['Underline'] }
                            If ($Blink) { $Parameters['Object'] += $ANSI['Blink'] }
                            If ($CrossedOut) { $Parameters['Object'] += $ANSI['CrossedOut'] }
                            If ($DoubleUnderline) { $Parameters['Object'] += $ANSI['DoubleUnderline'] }
                            If ($Overline) { $Parameters['Object'] += $ANSI['Overline'] }
                        }

                        # Set the foreground color to the default color based on the type of coloring for the line
                        If ($ANSI8) {
                            $Parameters['Object'] += "$esc[38;5;$($DefaultColor)m"
                        } ElseIf ($ANSI4) {
                            $Parameters['Object'] += "$esc[$($DefaultColor)m"
                        } Else {
                            $Parameters['ForegroundColor'] = $DefaultColor
                        }

                        # Set the foreground color based on the type of coloring for the line
                        If ($null -ne $BackGroundColor[$i]) {
                            If ($ANSI8) {
                                $Parameters['Object'] += "$esc[48;5;$($BackGroundColor[$i])m"
                            } ElseIf ($ANSI4) {
                                $Parameters['Object'] += "$esc[$($BackGroundColor[$i])m"
                            } Else {
                                $Parameters['BackgroundColor'] = $BackGroundColor[$i]
                            }
                        }

                        # Add the text for the text segment
                        $Parameters['Object'] += $Text[$i]

                        If ($ANSISupport) {
                            # Add the ANSI reset to stop the formatting after printing the text segment
                            $Parameters['Object'] += $ANSI['Reset']
                        }

                        Write-Host @Parameters
                    }
                }
            }
        }
        If ($NoNewLine -eq $true) {
            Write-Host -NoNewline 
        } Else {
            Write-Host 
        } # Support for no new line
        If ($LinesAfter -ne 0) {
            For ($i = 0; $i -lt $LinesAfter; $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        }  # Add empty line after
    }
    If ($Text.Count -and $LogFile) {
        If (!(Test-Path -Path "$LogPath")) {
            New-Item -ItemType 'Directory' -Path "$LogPath"
        }

        # LogFile is not a path, join the LogPath. This maintains compatibility with $LogFile while allowing a $LogName parameter.
        If ($LogFile -notmatch '[\\/]+') {
            If ($LogFile -notmatch '\.\w+$') {
                $LogFile += '.log'
            }
            $LogFilePath = Join-Path -Path $LogPath -ChildPath "$LogFile"
        } Else {
            $LogFilePath = $LogFile
        }

        # Save to file
        $TextToFile = ''
        For ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        $Saved = $False
        $Retry = 0
        Do {
            $Retry++
            try {
                $LogInfo = ''
                If ($LogTime) {
                    $LogInfo += "[$([datetime]::Now.ToString($DateTimeFormat))]"
                }

                If ($LogLevel.Length -gt 0 ) {
                    $LogInfo += "[$LogLevel]"
                }

                If (-not $LogInfo) {
                    "$TextToFile" | Out-File -FilePath $LogFilePath -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$False
                } Else {
                    "$LogInfo $TextToFile" | Out-File -FilePath $LogFilePath -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$False
                }
                $Saved = $true
            } Catch {
                If ($Saved -eq $False -and $Retry -eq $LogRetry) {
                    Write-Warning "Write-ColorEX - Couldn't write to log file $($_.Exception.Message). Tried ($Retry/$LogRetry))"
                } Else {
                    Write-Warning "Write-ColorEX - Couldn't write to log file $($_.Exception.Message). Retrying... ($Retry/$LogRetry)"
                }
            }
        } Until ($Saved -eq $true -or $Retry -ge $LogRetry)
    }
}

Function Update-WindowTitle {
    <#
    .SYNOPSIS
        Updates the PowerShell window title with current operation status
    .PARAMETER Action
        Current action being performed
    .PARAMETER TargetMailbox
        Target mailbox for the operation
    .PARAMETER UserMailbox
        User mailbox for permission operations
    .PARAMETER Step
        Current step in the process
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$Action,
        
        [Parameter(Mandatory = $false)]
        [String]$TargetMailbox,
        
        [Parameter(Mandatory = $false)]
        [String]$UserMailbox,
        
        [Parameter(Mandatory = $false)]
        [String]$Step
    )
    
    $TitleParts = [System.Collections.Generic.List[String]]::new()
    $TitleParts.Add('EXO Calendar Manager')
    
    If (-not [String]::IsNullOrEmpty($Action)) {
        $TitleParts.Add("| $Action")
    }
    
    If (-not [String]::IsNullOrEmpty($TargetMailbox)) {
        $TitleParts.Add("| Target: $TargetMailbox")
    }
    
    If (-not [String]::IsNullOrEmpty($UserMailbox)) {
        $TitleParts.Add("| User: $UserMailbox")
    }
    
    If (-not [String]::IsNullOrEmpty($Step)) {
        $TitleParts.Add("| $Step")
    }
    
    $Host.UI.RawUI.WindowTitle = $TitleParts -join ' '
}

Function Show-SampleNotificationEmail {
    <#
    .SYNOPSIS
        Displays a sample of the calendar sharing notification email that would be sent
    .PARAMETER FromDisplayName
        Display name of the calendar owner
    .PARAMETER FromEmailAddress
        Email address of the calendar owner
    .PARAMETER ToDisplayName
        Display name of the recipient
    .PARAMETER ToEmailAddress
        Email address of the recipient
    .PARAMETER AccessRights
        The access rights being granted
    .PARAMETER SharingPermissionFlags
        Additional sharing permission flags
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$FromDisplayName,
        
        [Parameter(Mandatory = $true)]
        [String]$FromEmailAddress,
        
        [Parameter(Mandatory = $true)]
        [String]$ToDisplayName,
        
        [Parameter(Mandatory = $true)]
        [String]$ToEmailAddress,
        
        [Parameter(Mandatory = $true)]
        [String]$AccessRights,
        
        [Parameter(Mandatory = $false)]
        [String]$SharingPermissionFlags = 'None'
    )
    
    # Create a sample notification email based on Exchange Online calendar sharing format
    $SampleWidth = [Math]::Max(80, [Math]::Max($FromDisplayName.Length + $FromEmailAddress.Length + 15, $ToDisplayName.Length + $ToEmailAddress.Length + 15))
    $SampleBorderTop = "┏" + ("━" * [Math]::Max(0, ($SampleWidth - 2))) + "┓"
    $SampleBorderMiddle = "┣" + ("━" * [Math]::Max(0, ($SampleWidth - 2))) + "┫"
    $SampleBorderBottom = "┗" + ("━" * [Math]::Max(0, ($SampleWidth - 2))) + "┛"
    
    Write-ColorEX -Text "  $SampleBorderTop" -Color LightBlue -ANSI8
    $EmailTitlePadding = [Math]::Max(0, $SampleWidth - 25 - 2)
    Write-ColorEX -Text "  ┃ ", "📧 Sample Email Preview", (" " * $EmailTitlePadding), " ┃" -Color LightBlue, LightBlue, None, LightBlue -Style None, @('Bold', 'Underline'), None, None -ANSI8
    Write-ColorEX -Text "  $SampleBorderMiddle" -Color LightBlue -ANSI8
    
    # Email header information
    $FromPadding = [Math]::Max(0, $SampleWidth - $FromDisplayName.Length - $FromEmailAddress.Length - 13)
    Write-ColorEX -Text "  ┃ ", "From: ", "$FromDisplayName", " <", "$FromEmailAddress", ">", (" " * $FromPadding), " ┃" -Color LightBlue, LightGray, LightCyan, LightGray, LightCyan, LightGray, None, LightBlue -ANSI8
    
    $ToPadding = [Math]::Max(0, $SampleWidth - $ToDisplayName.Length - $ToEmailAddress.Length - 11)
    Write-ColorEX -Text "  ┃ ", "To: ", "$ToDisplayName", " <", "$ToEmailAddress", ">", (" " * $ToPadding), " ┃" -Color LightBlue, LightGray, LightCyan, LightGray, LightCyan, LightGray, None, LightBlue -ANSI8
    
    $SubjectText = "Calendar sharing invitation from $FromDisplayName"
    $SubjectPadding = [Math]::Max(0, $SampleWidth - $SubjectText.Length - 13)
    Write-ColorEX -Text "  ┃ ", "Subject: ", "$SubjectText", (" " * $SubjectPadding), " ┃" -Color LightBlue, LightGray, LightGreen, None, LightBlue -ANSI8
    
    Write-ColorEX -Text "  $SampleBorderMiddle" -Color LightBlue -ANSI8
    
    # Email body content
    Write-ColorEX -Text "  ┃", (" " * ($SampleWidth - 2)), "┃" -Color LightBlue, None, LightBlue -ANSI8
    
    $BodyLine1 = "$FromDisplayName has shared their calendar with you."
    $Body1Padding = [Math]::Max(0, $SampleWidth - $BodyLine1.Length - 4)
    Write-ColorEX -Text "  ┃ ", "$BodyLine1", (" " * $Body1Padding), " ┃" -Color LightBlue, White, None, LightBlue -ANSI8
    
    Write-ColorEX -Text "  ┃", (" " * [Math]::Max(0, ($SampleWidth - 2))), "┃" -Color LightBlue, None, LightBlue -ANSI8
    
    # Permission level information
    $PermissionDescription = Switch ($AccessRights) {
        'Owner' { 'Full control of the calendar' }
        'PublishingEditor' { 'Create, read, modify, and delete items and folders' }
        'Editor' { 'Create, read, modify, and delete items' }
        'PublishingAuthor' { 'Create and read items and folders; modify and delete own items' }
        'Author' { 'Create and read items; modify and delete own items' }
        'NonEditingAuthor' { 'Create and read items; delete own items' }
        'Reviewer' { 'Read items only' }
        'Contributor' { 'Create items only' }
        'AvailabilityOnly' { 'View free/busy information only' }
        'LimitedDetails' { 'View free/busy and subject information' }
        Default { $AccessRights }
    }
    
    $PermLine = "Permission Level: $AccessRights ($PermissionDescription)"
    $PermPadding = $SampleWidth - $PermLine.Length - 4
    Write-ColorEX -Text "  ┃ ", "$PermLine", (" " * $PermPadding), " ┃" -Color LightBlue, LightYellow, None, LightBlue -ANSI8
    
    # Delegate information if applicable
    If ($SharingPermissionFlags -ne 'None') {
        $DelegateText = "Delegate Status: $SharingPermissionFlags"
        $DelegatePadding = $SampleWidth - $DelegateText.Length - 4
        Write-ColorEX -Text "  ┃ ", "$DelegateText", (" " * $DelegatePadding), " ┃" -Color LightBlue, LightMagenta, None, LightBlue -ANSI8
    }
    
    Write-ColorEX -Text "  ┃", (" " * ($SampleWidth - 2)), "┃" -Color LightBlue, None, LightBlue -ANSI8
    
    $AcceptText = "To start using this shared calendar, click Accept below."
    $AcceptPadding = $SampleWidth - $AcceptText.Length - 4
    Write-ColorEX -Text "  ┃ ", "$AcceptText", (" " * $AcceptPadding), " ┃" -Color LightBlue, White, None, LightBlue -ANSI8
    
    Write-ColorEX -Text "  ┃", (" " * ($SampleWidth - 2)), "┃" -Color LightBlue, None, LightBlue -ANSI8
    
    # Accept button representation
    $ButtonText = "[ Accept Calendar Sharing Invitation ]"
    $ButtonPadding = ($SampleWidth - $ButtonText.Length - 2) / 2
    $ButtonPaddingLeft = [Math]::Floor($ButtonPadding)
    $ButtonPaddingRight = [Math]::Ceiling($ButtonPadding)
    Write-ColorEX -Text "  ┃", (" " * $ButtonPaddingLeft), "$ButtonText", (" " * $ButtonPaddingRight), "┃" -Color LightBlue, None, LightGreen, None, LightBlue -BackGroundColor None, None, DarkGreen, None, None -Style None, None, Bold, None, None -ANSI8
    
    Write-ColorEX -Text "  ┃", (" " * ($SampleWidth - 2)), "┃" -Color LightBlue, None, LightBlue -ANSI8
    
    $FooterText = "This is a normal calendar sharing invitation from Exchange Online."
    $FooterPadding = $SampleWidth - $FooterText.Length - 4
    Write-ColorEX -Text "  ┃ ", "$FooterText", (" " * $FooterPadding), " ┃" -Color LightBlue, DarkGray, None, LightBlue -ANSI8
    
    Write-ColorEX -Text "  $SampleBorderBottom" -Color LightBlue -ANSI8
    Write-ColorEX '' -LinesBefore 1
}

Function Show-StatusBar {
    <#
    .SYNOPSIS
        Displays a status bar at the top of the console with current operation details
    #>
    [CmdletBinding()]
    Param()
    
    If ($Script:CurrentAction -or $Script:CurrentTargetMailbox -or $Script:CurrentUserMailbox) {
        $WindowWidth = $Host.UI.RawUI.BufferSize.Width
        $StatusLine = [System.Collections.Generic.List[String]]::new()
        
        If ($Script:CurrentAction) {
            $StatusLine.Add("Action: $($Script:CurrentAction)")
        }
        
        If ($Script:CurrentTargetMailbox) {
            $StatusLine.Add("Target: $($Script:CurrentTargetMailbox)")
        }
        
        If ($Script:CurrentUserMailbox) {
            $StatusLine.Add("User: $($Script:CurrentUserMailbox)")
        }
        
        If ($Script:CurrentStep) {
            $StatusLine.Add("Step: $($Script:CurrentStep)")
        }
        
        $StatusText = $StatusLine -join ' | '
        $PaddingNeeded = [Math]::Max(0, $WindowWidth - $StatusText.Length - 4)
        $PaddedStatus = " $StatusText" + (' ' * $PaddingNeeded) + ' '
        
        Write-ColorEX -Text $PaddedStatus -Color White -BackGroundColor DarkBlue -ANSI8
        Write-ColorEX '' # Empty line for separation
    }
}

Function Reset-OperationState {
    <#
    .SYNOPSIS
        Resets the current operation state and updates title
    #>
    [CmdletBinding()]
    Param()
    
    $Script:CurrentAction = $null
    $Script:CurrentTargetMailbox = $null
    $Script:CurrentUserMailbox = $null
    $Script:CurrentStep = $null
    Update-WindowTitle
}

Function Set-OperationState {
    <#
    .SYNOPSIS
        Sets the current operation state and updates title and status
    .PARAMETER Action
        Current action being performed
    .PARAMETER TargetMailbox
        Target mailbox for the operation
    .PARAMETER UserMailbox
        User mailbox for permission operations
    .PARAMETER Step
        Current step in the process
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$Action,
        
        [Parameter(Mandatory = $false)]
        [String]$TargetMailbox,
        
        [Parameter(Mandatory = $false)]
        [String]$UserMailbox,
        
        [Parameter(Mandatory = $false)]
        [String]$Step
    )
    
    If ($Action) { $Script:CurrentAction = $Action }
    If ($TargetMailbox) { $Script:CurrentTargetMailbox = $TargetMailbox }
    If ($UserMailbox) { $Script:CurrentUserMailbox = $UserMailbox }
    If ($Step) { $Script:CurrentStep = $Step }
    
    Update-WindowTitle -Action $Script:CurrentAction -TargetMailbox $Script:CurrentTargetMailbox -UserMailbox $Script:CurrentUserMailbox -Step $Script:CurrentStep
}

Function Show-InteractiveMenu {
    <#
    .SYNOPSIS
        Displays an advanced interactive menu with arrow key navigation only and Unicode borders
    .PARAMETER Title
        The menu title
    .PARAMETER Options
        Array of menu options
    .PARAMETER AllowBack
        Whether to show a "Back" option
    .PARAMETER AllowQuit
        Whether to show a "Quit" option
    .PARAMETER ShowStatusBar
        Whether to display the status bar
    .PARAMETER ShowSampleEmail
        Whether to show sample notification email
    .PARAMETER SampleEmailParams
        Parameters for sample email display
    .PARAMETER SummaryContent
        Array of summary content to display before menu
    .OUTPUTS
        PSCustomObject with Selected (index), Action ('Select', 'Back', 'Quit', 'Cancel')
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$Title,
        
        [Parameter(Mandatory = $true)]
        [String[]]$Options,
        
        [Parameter(Mandatory = $false)]
        [Boolean]$AllowBack = $false,
        
        [Parameter(Mandatory = $false)]
        [Boolean]$AllowQuit = $false,
        
        [Parameter(Mandatory = $false)]
        [Boolean]$ShowStatusBar = $true,

        [Parameter(Mandatory = $false)]
        [Boolean]$ShowSampleEmail = $false,
        
        [Parameter(Mandatory = $false)]
        [Hashtable]$SampleEmailParams = @{},
        
        [Parameter(Mandatory = $false)]
        [Array]$SummaryContent = @()
    )

    [Console]::CursorVisible = $false
    
    # Prepare menu options with navigation
    $MenuOptions = [System.Collections.Generic.List[String]]::new()
    $MenuOptions.AddRange($Options)
    
    If ($AllowBack) {
        $MenuOptions.Add('🔙 Back')
    }
    
    If ($AllowQuit) {
        $MenuOptions.Add('❌ Quit')
    }
    
    $CurrentSelection = 0
    $MaxSelection = $MenuOptions.Count - 1
    
    Do {
        Clear-Host

        # Show status bar if enabled and there's status to show
        If ($ShowStatusBar) {
            Show-StatusBar
        }
        
        # Show sample email if requested
        If ($ShowSampleEmail -and $SampleEmailParams.Count -gt 0) {
            Show-SampleNotificationEmail @SampleEmailParams
        }
        
        # Show summary content if provided
        If ($SummaryContent.Count -gt 0) {
            ForEach ($SummaryLine in $SummaryContent) {
                & $SummaryLine
            }
        }
        
        # Calculate menu width
        $MenuWidth = [Math]::Max(60, $Title.Length + 10)
        ForEach ($Option in $MenuOptions) {
            $MenuWidth = [Math]::Max($MenuWidth, $Option.Length + 10)
        }
        
        $BorderTop = "┏" + ("━" * ($MenuWidth - 2)) + "┓"
        $BorderMiddle = "┣" + ("━" * ($MenuWidth - 2)) + "┫"
        $BorderBottom = "┗" + ("━" * ($MenuWidth - 2)) + "┛"
        
        # Menu header with Unicode border
        Write-ColorEX -Text "  $BorderTop" -Color Cyan -ANSI8
        $TitlePadding = $MenuWidth - $Title.Length - 4
        Write-ColorEX -Text "  ┃ ", "$Title", (" " * $TitlePadding), " ┃" -Color Cyan, LightCyan, None, Cyan -Style None, @('Bold', 'Underline'), None, None -ANSI8
        Write-ColorEX -Text "  $BorderMiddle" -Color Cyan -ANSI8
        
        # Navigation instructions inside border
        $NavText = "↑↓ or ←→ arrows, Enter to select, Esc to cancel"
        $NavPadding = $MenuWidth - $NavText.Length - 4
        Write-ColorEX -Text "  ┃ ", "$NavText", (" " * $NavPadding), " ┃" -Color Cyan, Yellow, None, Cyan -ANSI8
        
        Write-ColorEX -Text "  $BorderMiddle" -Color Cyan -ANSI8
        
        # Display menu options with highlighting inside border
        For ($i = 0; $i -lt $MenuOptions.Count; $i++) {
            $IsBack = $MenuOptions[$i] -like '*Back*'
            $IsQuit = $MenuOptions[$i] -like '*Quit*'
            
            $OptionText = $MenuOptions[$i]
            $Padding = $MenuWidth - $OptionText.Length - 6  # Account for border and arrow
            
            If ($i -eq $CurrentSelection) {
                If ($IsBack) {
                    Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * $Padding), " ┃" -Color Cyan, LightMagenta, Yellow, None, Cyan -BackGroundColor None, None, DarkMagenta, None, None -Style None, Bold, Bold, None, None -ANSI8
                } ElseIf ($IsQuit) {
                    Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * ($Padding - 1)), " ┃" -Color Cyan, LightRed, White, None, Cyan -BackGroundColor None, None, DarkRed, None, None -Style None, Bold, Bold, None, None -ANSI8
                } Else {
                    If ($OptionText -match "🗑️") {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * ($Padding + 1)), " ┃" -Color Cyan, LightBlue, White, None, Cyan -BackGroundColor None, None, DarkBlue, None, None -Style None, Bold, Bold, None, None -ANSI8
                    } ElseIf ($OptionText -match "✅" -or $OptionText -match "❌") {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * ($Padding - 1)), " ┃" -Color Cyan, LightBlue, White, None, Cyan -BackGroundColor None, None, DarkBlue, None, None -Style None, Bold, Bold, None, None -ANSI8
                    } Else {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * $Padding), " ┃" -Color Cyan, LightBlue, White, None, Cyan -BackGroundColor None, None, DarkBlue, None, None -Style None, Bold, Bold, None, None -ANSI8
                    }
                }
            } Else {
                If ($IsBack) {
                    Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * $Padding), " ┃" -Color Cyan, LightMagenta, None, Cyan -ANSI8
                } ElseIf ($IsQuit) {
                    Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * ($Padding - 1)), " ┃" -Color Cyan, LightRed, None, Cyan -ANSI8
                } Else {
                    If ($OptionText -match "🗑️") {
                        Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * ($Padding + 1)), " ┃" -Color Cyan, LightGray, None, Cyan -ANSI8
                    } ElseIf ($OptionText -match "✅" -or $OptionText -match "❌") {
                        Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * ($Padding - 1)), " ┃" -Color Cyan, LightGray, None, Cyan -ANSI8
                    } Else {
                        Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * $Padding), " ┃" -Color Cyan, LightGray, None, Cyan -ANSI8
                    }
                }
            }
        }
        
        # Close the border
        Write-ColorEX -Text "  $BorderBottom" -Color Cyan -ANSI8
        Write-ColorEX ''
        
        # Wait for key input with proper error handling
        Try {
            $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            
            Switch ($Key.VirtualKeyCode) {
                38 { # Up Arrow
                    $CurrentSelection = If ($CurrentSelection -eq 0) { $MaxSelection } Else { $CurrentSelection - 1 }
                }
                40 { # Down Arrow  
                    $CurrentSelection = If ($CurrentSelection -eq $MaxSelection) { 0 } Else { $CurrentSelection + 1 }
                }
                37 { # Left Arrow (same as up)
                    $CurrentSelection = If ($CurrentSelection -eq 0) { $MaxSelection } Else { $CurrentSelection - 1 }
                }
                39 { # Right Arrow (same as down)
                    $CurrentSelection = If ($CurrentSelection -eq $MaxSelection) { 0 } Else { $CurrentSelection + 1 }
                }
                13 { # Enter
                    $SelectedOption = $MenuOptions[$CurrentSelection]
                    If ($SelectedOption -like '*Back*') {
                        [Console]::CursorVisible = $True; Return [PSCustomObject]@{ Selected = -1; Action = 'Back' }
                    } ElseIf ($SelectedOption -like '*Quit*') {
                        [Console]::CursorVisible = $True; Return [PSCustomObject]@{ Selected = -1; Action = 'Quit' }
                    } Else {
                        [Console]::CursorVisible = $True; Return [PSCustomObject]@{ Selected = $CurrentSelection; Action = 'Select' }
                    }
                }
                27 { # Escape
                    [Console]::CursorVisible = $True; Return [PSCustomObject]@{ Selected = -1; Action = 'Cancel' }
                }
                Default {
                    # Removed number key handling - only arrow keys work now
                }
            }
        } Catch {
            # Handle any key reading errors gracefully
            Start-Sleep -Milliseconds 50
        }
    } While ($true)
}

Function Get-ValidatedInput {
    <#
    .SYNOPSIS
        Gets validated user input with confirmation and cancel options using arrow navigation
    .PARAMETER Prompt
        The prompt message
    .PARAMETER ValidationType
        Type of validation to perform (Email, None)
    .PARAMETER AllowEmpty
        Whether empty input is allowed
    .OUTPUTS
        PSCustomObject with Value (string), Action ('Confirm', 'Cancel')
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$Prompt,

        [Parameter(Mandatory = $false)]
        [String]$PromptTitle = 'Input Required',
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Email', 'None')]
        [String]$ValidationType = 'None',
        
        [Parameter(Mandatory = $false)]
        [Boolean]$AllowEmpty = $false
    )
    
    Do {
        Clear-Host
        Show-StatusBar
        
        $InputWidth = [Math]::Max(60, $Prompt.Length + 10)
        $BorderTop = "┏" + ("━" * ($InputWidth - 2)) + "┓"
        $BorderBottom = "┗" + ("━" * ($InputWidth - 2)) + "┛"
        
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text "  $BorderTop" -Color LightYellow -ANSI8
        $TitlePadding = $InputWidth - $PromptTitle.Length - 4  # "Input Required" is 15 chars
        Write-ColorEX -Text "  ┃ ", "$PromptTitle", (" " * $TitlePadding), " ┃" -Color LightYellow, LightYellow, None, LightYellow -Style None, @('Bold', 'Underline'), None, None -ANSI8
        Write-ColorEX -Text "  ┃", (" " * ($InputWidth - 2)), "┃" -Color LightYellow, None, LightYellow -ANSI8
        $PromptPadding = $InputWidth - $Prompt.Length - 4
        Write-ColorEX -Text "  ┃ ", "$Prompt", (" " * $PromptPadding), " ┃" -Color LightYellow, LightCyan, None, LightYellow -ANSI8
        Write-ColorEX -Text "  $BorderBottom" -Color LightYellow -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text "  $PromptTitle → " -Color LightGray -NoNewLine -ANSI8
        
        $UserInput = Read-Host
        
        # Check for empty input
        If ([String]::IsNullOrWhiteSpace($UserInput)) {
            If ($AllowEmpty) {
                $UserInput = $UserInput.Trim()
            } Else {
                Write-ColorEX -Text '[', '⚠️ WARNING', '] Input cannot be empty.' -Color White, Orange, White -ANSI8
                Write-ColorEX -Text 'Press any key to try again...' -Color Yellow -ANSI8
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                Continue
            }
        } Else {
            $UserInput = $UserInput.Trim()
        }
        
        # Validate based on type
        $IsValid = $true
        Switch ($ValidationType) {
            'Email' {
                If (-not (Test-EmailFormat -EmailAddress $UserInput)) {
                    Write-ColorEX -Text '[', '⚠️ WARNING', '] Invalid email format. Please enter a valid email address.' -Color White, Orange, White -ANSI8
                    $IsValid = $false
                }
            }
        }
        
        If (-not $IsValid) {
            Write-ColorEX -Text 'Press any key to try again...' -Color Yellow -ANSI8
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            Continue
        }
        
        # Custom confirmation menu that shows the input value
        # Hide the cursor
        [Console]::CursorVisible = $False

        $CurrentSelection = 0
        $ConfirmOptions = @(
            '✅ Confirm - Use this value',
            '🔄 Retry - Enter a different value',
            '❌ Cancel - Return to previous menu'
        )
        $MaxSelection = $ConfirmOptions.Count - 1
        
        Do {
            Clear-Host
            Show-StatusBar
            
            $ConfirmWidth = [Math]::Max(60, $UserInput.Length + 20)
            $ConfirmBorderTop = "┏" + ("━" * ($ConfirmWidth - 2)) + "┓"
            $ConfirmBorderMiddle = "┣" + ("━" * ($ConfirmWidth - 2)) + "┫"
            $ConfirmBorderBottom = "┗" + ("━" * ($ConfirmWidth - 2)) + "┛"
            
            # Show the input being confirmed
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text "  $ConfirmBorderTop" -Color LightYellow -ANSI8
            $ConfirmTitlePadding = $ConfirmWidth - 18 - 4  # "Confirm Your Input" is 18 chars
            Write-ColorEX -Text "  ┃ ", 'Confirm Your Input', (" " * $ConfirmTitlePadding), " ┃" -Color LightYellow, LightYellow, None, LightYellow -Style None, @('Bold', 'Underline'), None, None -ANSI8
            Write-ColorEX -Text "  $ConfirmBorderMiddle" -Color LightYellow -ANSI8
            $InputPadding = $ConfirmWidth - $UserInput.Length - 19  # "Input Value: " + quotes is ~15 chars
            Write-ColorEX -Text "  ┃ ", 'Input Value: ', "'$UserInput'", (" " * $InputPadding), " ┃" -Color LightYellow, LightGray, LightGreen, None, LightYellow -ANSI8
            Write-ColorEX -Text "  $ConfirmBorderMiddle" -Color LightYellow -ANSI8
            
            # Navigation instructions
            $NavText = "↑↓ or ←→ arrows, Enter to select, Esc to cancel"
            $NavPadding = $ConfirmWidth - $NavText.Length - 4
            Write-ColorEX -Text "  ┃ ", "$NavText", (" " * $NavPadding), " ┃" -Color LightYellow, Yellow, None, LightYellow -ANSI8
            Write-ColorEX -Text "  $ConfirmBorderMiddle" -Color LightYellow -ANSI8
            
            # Display confirmation options with highlighting
            For ($i = 0; $i -lt $ConfirmOptions.Count; $i++) {
                $OptionText = $ConfirmOptions[$i]
                $Padding = $ConfirmWidth - $OptionText.Length - 7  # Account for border and arrow
                
                If ($i -eq $CurrentSelection) {
                    If ($OptionText -like "*Confirm*") {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * $Padding), " ┃" -Color LightYellow, LightGreen, White, None, LightYellow -BackGroundColor None, None, DarkGreen, None, None -Style None, Bold, Bold, None, None -ANSI8
                    } ElseIf ($OptionText -like "*Retry*") {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * ($Padding + 1)), " ┃" -Color LightYellow, LightBlue, White, None, LightYellow -BackGroundColor None, None, DarkBlue, None, None -Style None, Bold, Bold, None, None -ANSI8
                    } ElseIf ($OptionText -like "*Cancel*") {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * $Padding), " ┃" -Color LightYellow, LightRed, White, None, LightYellow -BackGroundColor None, None, DarkRed, None, None -Style None, Bold, Bold, None, None -ANSI8
                    } Else {
                        Write-ColorEX -Text "  ┃ ", "→ ", "$OptionText", (" " * $Padding), " ┃" -Color LightYellow, LightGray, White, None, LightYellow -BackGroundColor None, None, DarkGray, None, None -Style None, Bold, Bold, None, None -ANSI8
                    }
                } Else {
                    If ($OptionText -like "*Retry*") {
                        Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * ($Padding + 1)), " ┃" -Color LightYellow, LightGray, None, LightYellow -ANSI8
                    } Else {
                        Write-ColorEX -Text "  ┃   ", "$OptionText", (" " * $Padding), " ┃" -Color LightYellow, LightGray, None, LightYellow -ANSI8
                    }
                    
                }
            }
            
            Write-ColorEX -Text "  $ConfirmBorderBottom" -Color LightYellow -ANSI8
            Write-ColorEX ''
            
            # Wait for key input with proper error handling
            Try {
                $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                
                Switch ($Key.VirtualKeyCode) {
                    38 { # Up Arrow
                        $CurrentSelection = If ($CurrentSelection -eq 0) { $MaxSelection } Else { $CurrentSelection - 1 }
                    }
                    40 { # Down Arrow  
                        $CurrentSelection = If ($CurrentSelection -eq $MaxSelection) { 0 } Else { $CurrentSelection + 1 }
                    }
                    37 { # Left Arrow (same as up)
                        $CurrentSelection = If ($CurrentSelection -eq 0) { $MaxSelection } Else { $CurrentSelection - 1 }
                    }
                    39 { # Right Arrow (same as down)
                        $CurrentSelection = If ($CurrentSelection -eq $MaxSelection) { 0 } Else { $CurrentSelection + 1 }
                    }
                    13 { # Enter
                        Switch ($CurrentSelection) {
                            0 { # Confirm
                                [Console]::CursorVisible = $False; Write-ColorEX ''; Return [PSCustomObject]@{ Value = $UserInput; Action = 'Confirm' }
                            }
                            1 { # Retry
                                $RetrySelected = $true
                                Break
                            }
                            2 { # Cancel
                                [Console]::CursorVisible = $False; Write-ColorEX ''; Return [PSCustomObject]@{ Value = $null; Action = 'Cancel' }
                            }
                        }
                    }
                    27 { # Escape
                        [Console]::CursorVisible = $False; Write-ColorEX ''; Return [PSCustomObject]@{ Value = $null; Action = 'Cancel' }
                    }
                    Default {
                        # Ignore other keys
                    }
                }
            } Catch {
                # Handle any key reading errors gracefully
                Start-Sleep -Milliseconds 50
            }
        } While (-not $RetrySelected)
        
        # Reset retry flag for next iteration
        $RetrySelected = $false
        
    } While ($true)
}

Function Test-EmailFormat {
    <#
    .SYNOPSIS
        Validates email address format using both regex and MailAddress class
    .PARAMETER EmailAddress
        The email address to validate
    .OUTPUTS
        Boolean indicating if the email format is valid
    #>
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$EmailAddress
    )
    
    # First check with regex for basic format validation
    $EmailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    If (-not ($EmailAddress -match $EmailRegex)) {
        Return $false
    }
    
    # Second validation using .NET MailAddress class for comprehensive validation
    Try {
        $null = [System.Net.Mail.MailAddress]$EmailAddress
        Return $true
    } Catch {
        Return $false
    }
}

Function Test-ModuleInstalled {
    <#
    .SYNOPSIS
        Checks if the Exchange Online Management module is installed and up to date
    .OUTPUTS
        Boolean indicating if the module is properly installed
    #>
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param()
    
    Try {
        # Check both PowerShellGet and PSResourceGet installations
        $InstalledModule = $null
        
        # Try PSResourceGet first (newer method)
        Try {
            $InstalledModule = Get-InstalledPSResource -Name $Script:ModuleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        } Catch {
            # Fallback to PowerShellGet
            $InstalledModule = Get-InstalledModule -Name $Script:ModuleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        }
        
        If ($null -eq $InstalledModule) {
            Return $false
        }
        
        # Version checking logic
        $CurrentVersion = [System.Version]$InstalledModule.Version
        $RequiredVersion = [System.Version]$Script:RequiredModuleVersion
        
        If ($CurrentVersion -lt $RequiredVersion) {
            Write-ColorEX -Text '[', '⚠️ WARNING', '] Current module version ', "$($CurrentVersion)", ' is below required version ', "$($RequiredVersion)" -Color White, Orange, White, Yellow, White, Green -ANSI8
            Return $false
        } Else {
            Write-ColorEX -Text '[', '✅ SUCCESS', "] Exchange Online Management module minimum version check passed: ", "$($CurrentVersion)" -Color White, Green, White, LightGreen -ANSI8
        }
        
        Return $true
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error checking module installation: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Return $false
    }
}

Function Install-ExchangeModule {
    <#
    .SYNOPSIS
        Installs or updates the Exchange Online Management module using modern PSResourceGet
    .DESCRIPTION
        Uses PSResourceGet instead of PowerShellGet for better compatibility with v3.8.0
    #>
    [CmdletBinding()]
    Param()
    
    Try {
        Write-ColorEX -Text '[', 'ℹ️ INFO', '] Installing/Updating Exchange Online Management module...' -Color White, Cyan, White -ANSI8
        
        # IMPROVED: Check if PSResourceGet is available (recommended for v3.8.0)
        $PSResourceGetAvailable = Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable
        
        If ($PSResourceGetAvailable) {
            Write-ColorEX -Text '[', 'ℹ️ INFO', '] Using PSResourceGet for module installation...' -Color White, Cyan, White -ANSI8
            
            # UPDATED: Use Install-PSResource (newer method recommended by Microsoft)
            Install-PSResource -Name $Script:ModuleName -Scope CurrentUser -TrustRepository -Force -ErrorAction Stop
            
            Write-ColorEX -Text '[', '✅ SUCCESS', '] Exchange Online Management module installed successfully using PSResourceGet' -Color White, Green, White -ANSI8
        } Else {
            Write-ColorEX -Text '[', 'ℹ️ INFO', '] Using PowerShellGet for module installation...' -Color White, Cyan, White -ANSI8
            
            # Fallback to traditional method
            $Repository = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
            If ($Repository.InstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
            }
            
            Install-Module -Name $Script:ModuleName -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
            Write-ColorEX -Text '[', '✅ SUCCESS', '] Exchange Online Management module installed successfully using PowerShellGet' -Color White, Green, White -ANSI8
        }
        
        # Import the module with better error handling
        Import-Module -Name $Script:ModuleName -RequiredVersion $Script:RequiredModuleVersion -Force -Global -ErrorAction Stop
        Write-ColorEX -Text '[', '✅ SUCCESS', '] Module imported successfully' -Color White, Green, White -ANSI8
        
        Write-ColorEX -Text 'Press any key to continue...' -Color DarkYellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Failed to install or import Exchange Online Management module: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Throw "Module installation failed: $($_.Exception.Message)"
    }
}

Function Connect-ExchangeOnlineService {
    <#
    .SYNOPSIS
        Establishes connection to Exchange Online using modern authentication
    .PARAMETER UserPrincipalName
        The admin UPN for connection
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$UserPrincipalName
    )
    
    Try {
        Write-ColorEX -Text '[', 'ℹ️ INFO', '] Connecting to Exchange Online as ', "$UserPrincipalName", '...' -Color White, Cyan, White, Yellow, White -ANSI8
        
        # UPDATED: Enhanced connection parameters for v3.8.0
        $ConnectParams = @{
            UserPrincipalName = $UserPrincipalName
            ShowProgress = $true
            ErrorAction = 'Stop'
            ShowBanner = $false  # IMPROVED: Suppress banner for cleaner output
        }
        
        # IMPROVED: Add CommandName parameter to reduce memory footprint for calendar operations
        $CalendarCmdlets = @(
            'Get-Mailbox',
            'Get-MailboxFolderStatistics', 
            'Get-MailboxFolderPermission',
            'Add-MailboxFolderPermission',
            'Set-MailboxFolderPermission',
            'Remove-MailboxFolderPermission',
            'Get-ConnectionInformation',
            'Disconnect-ExchangeOnline'
        )
        $ConnectParams.CommandName = $CalendarCmdlets
        
        # IMPROVED: Add SkipLoadingFormatData for better performance in Windows services
        If ($Host.Name -eq 'ConsoleHost') {
            $ConnectParams.SkipLoadingFormatData = $true
        }
        
        # Connect using optimized parameters
        Connect-ExchangeOnline @ConnectParams
        
        $Script:IsConnected = $true
        Write-ColorEX -Text '[', '✅ SUCCESS', '] Successfully connected to Exchange Online' -Color White, Green, White -ANSI8
        
        # UPDATED: Verify connection using Get-ConnectionInformation (replaces old session checks)
        $TestResult = Get-ConnectionInformation -ErrorAction SilentlyContinue
        If ($null -ne $TestResult) {
            Write-ColorEX -Text '[', '✅ SUCCESS', '] Connection verified ─ Tenant: ', "$($TestResult.TenantId)" -Color White, Green, White, LightGreen -ANSI8
        }
        
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Failed to connect to Exchange Online: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        $Script:IsConnected = $false
        Throw "Exchange Online connection failed: $($_.Exception.Message)"
    }
}

Function Test-MailboxExists {
    <#
    .SYNOPSIS
        Verifies if a mailbox exists in Exchange Online
    .PARAMETER Identity
        The mailbox identity (UPN, alias, or display name)
    .OUTPUTS
        PSCustomObject with mailbox information if found, $null if not found
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$Identity
    )
    
    Try {
        # Try to get mailbox information
        $Mailbox = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
        If ($null -ne $Mailbox) {
            Return [PSCustomObject]@{
                DisplayName = $Mailbox.DisplayName
                PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                RecipientTypeDetails = $Mailbox.RecipientTypeDetails
                Identity = $Mailbox.Identity
                Alias = $Mailbox.Alias
            }
        }
        Return $null
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error checking mailbox existence: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Write-ColorEX -Text 'Press any key to continue...' -Color DarkYellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return $null
    }
}

Function Get-CalendarFolderPath {
    <#
    .SYNOPSIS
        Gets the calendar folder path for a specific calendar
    .PARAMETER MailboxIdentity
        The mailbox identity
    .PARAMETER CalendarIdentity
        The specific calendar identity (optional, defaults to main calendar)
    .OUTPUTS
        String containing the calendar folder path
    #>
    [CmdletBinding()]
    [OutputType([String])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity,
        
        [Parameter(Mandatory = $false)]
        [String]$CalendarIdentity = $null
    )
    
    Try {
        # If specific calendar identity provided, use it
        If (-not [String]::IsNullOrEmpty($CalendarIdentity)) {
            Return $CalendarIdentity
        }
        
        # Otherwise, get the default calendar folder path
        $CalendarFolder = Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope Calendar -ErrorAction Stop | 
                         Where-Object { $_.FolderType -eq 'Calendar' } | 
                         Select-Object -First 1
        
        If ($null -ne $CalendarFolder) {
            $FolderPath = $CalendarFolder.FolderPath -replace '^/', ''
            Return "$($MailboxIdentity):\$FolderPath"
        } Else {
            # Fallback to default Calendar path
            Return "$($MailboxIdentity):\Calendar"
        }
    } Catch {
        Write-ColorEX -Text '[', '⚠️ WARNING', '] Could not determine calendar folder path, using default: ', "$($_.Exception.Message)" -Color White, Orange, White, Yellow -ANSI8
        Return "$($MailboxIdentity):\Calendar"
    }
}

Function Get-AvailableCalendars {
    <#
    .SYNOPSIS
        Gets all available calendars for a mailbox
    .PARAMETER MailboxIdentity
        The mailbox identity
    .OUTPUTS
        Array of calendar objects with Name and FolderPath
    #>
    [CmdletBinding()]
    [OutputType([Array])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity
    )
    
    Try {
        # Get all calendar folders
        $CalendarFolders = Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope Calendar -ErrorAction Stop
        
        # Create calendar objects with friendly names
        $Calendars = @()
        ForEach ($Folder in $CalendarFolders) {
            # Extract calendar name from folder path
            $CalendarName = If ($Folder.FolderType -eq 'Calendar') {
                'Calendar (Default)'
            } Else {
                $Folder.Name
            }
            
            $Calendars += [PSCustomObject]@{
                Name = $CalendarName
                FolderPath = $Folder.FolderPath
                FolderType = $Folder.FolderType
                ItemCount = $Folder.ItemsInFolder
                Size = $Folder.FolderSize
                Identity = "$MailboxIdentity`:$($Folder.FolderPath -replace '/', '\')"
            }
        }
        
        Return $Calendars | Sort-Object { If ($_.FolderType -eq 'Calendar') { 0 } Else { 1 } }, Name
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error retrieving calendars: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Return @()
    }
}

Function Select-Calendar {
    <#
    .SYNOPSIS
        Shows calendar selection menu or auto-selects if only default calendar exists
    .PARAMETER Calendars
        Array of available calendars
    .PARAMETER MailboxDisplayName
        Display name of the mailbox for the menu title
    .OUTPUTS
        Selected calendar object or $null if cancelled
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param(
        [Parameter(Mandatory = $true)]
        [Array]$Calendars,
        
        [Parameter(Mandatory = $true)]
        [String]$MailboxDisplayName
    )
    
    # If only default calendar exists, auto-select it
    If ($Calendars.Count -eq 1 -and $Calendars[0].FolderType -eq 'Calendar') {
        Return $Calendars[0]
    }
    
    # Multiple calendars found, show selection menu
    $CalendarOptions = @()
    ForEach ($Calendar in $Calendars) {
        $CalendarOptions += "$($Calendar.Name) ($($Calendar.ItemCount) items)"
    }
    
    $MenuResult = Show-InteractiveMenu -Title "Select Calendar for $MailboxDisplayName" -Options $CalendarOptions -AllowBack $true
    
    If ($MenuResult.Action -eq 'Select') {
        Return $Calendars[$MenuResult.Selected]
    }
    
    Return $null
}

Function Get-CalendarPermissions {
    <#
    .SYNOPSIS
        Retrieves current calendar permissions for a mailbox
    .PARAMETER MailboxIdentity
        The mailbox identity
    .PARAMETER CalendarPath
        The specific calendar folder path
    .OUTPUTS
        Array of permission objects
    #>
    [CmdletBinding()]
    [OutputType([Array])]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$CalendarPath
    )
    
    Try {
        $Permissions = Get-MailboxFolderPermission -Identity $CalendarPath -ErrorAction Stop
        
        Return $Permissions | Select-Object @{
            Name = 'User'
            Expression = { $_.User.DisplayName }
        }, @{
            Name = 'AccessRights'
            Expression = { $_.AccessRights -join ', ' }
        }, @{
            Name = 'SharingPermissionFlags'
            Expression = { $_.SharingPermissionFlags -join ', ' }
        }
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error retrieving calendar permissions: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Return @()
    }
}

Function Add-CalendarPermission {
    <#
    .SYNOPSIS
        Adds calendar permission for a user
    .PARAMETER MailboxIdentity
        The target mailbox identity
    .PARAMETER CalendarPath
        The specific calendar folder path
    .PARAMETER UserIdentity  
        The user to grant permissions to
    .PARAMETER AccessRights
        The access rights to grant
    .PARAMETER SendNotificationToUser
        Whether to send email notification to the user
    .PARAMETER SharingPermissionFlags
        Additional sharing permission flags
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$CalendarPath,
        
        [Parameter(Mandatory = $true)]
        [String]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$AccessRights,

        [Parameter(Mandatory = $true)]
        [Bool]$SendNotificationToUser,
        
        [Parameter(Mandatory = $false)]
        [String]$SharingPermissionFlags = 'None'
    )
    
    Try {
        $AddParams = @{
            Identity = $CalendarPath
            User = $UserIdentity
            AccessRights = $AccessRights
            SendNotificationToUser = $SendNotificationToUser
            ErrorAction = 'Stop'
        }
        
        # Add sharing permission flags if specified and access rights is Editor
        If ($AccessRights -eq 'Editor' -and $SharingPermissionFlags -ne 'None') {
            $AddParams.SharingPermissionFlags = $SharingPermissionFlags
        }
        
        Write-Progress -Activity 'Adding calendar permissions' -Status 'Progress ->' -PercentComplete 50
        Add-MailboxFolderPermission @AddParams
        Write-Progress -Activity 'Adding calendar permissions' -Status 'Progress ->' -Completed
        
        Write-ColorEX -Text '[', '✅ SUCCESS', '] Successfully added ', "$AccessRights", ' permission for ', "$UserIdentity", ' on ', "$MailboxIdentity", ' calendar' -Color White, Green, White, LightGreen, White, LightBlue, White, LightBlue, White -ANSI8
        
        If ($SendNotificationToUser) {
            Write-ColorEX -Text '[', 'ℹ️ INFO', '] Email notification sent to ', "$UserIdentity", ' about calendar sharing' -Color White, Cyan, White, LightBlue, White -ANSI8
        }
        
    } Catch {
        If ($_.Exception.Message -like '*already has permission*') {
            Write-ColorEX -Text '[', '⚠️ WARNING', '] User ', "$UserIdentity", ' already has permissions. Use ', "'Modify Permission'", ' instead.' -Color White, Orange, White, Yellow, White, LightYellow, White -ANSI8
        } Else {
            Write-ColorEX -Text '[', '❌ ERROR', '] Error adding calendar permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
            Throw
        }
    }
}

Function Set-CalendarPermission {
    <#
    .SYNOPSIS
        Modifies existing calendar permission for a user
    .PARAMETER MailboxIdentity
        The target mailbox identity
    .PARAMETER CalendarPath
        The specific calendar folder path
    .PARAMETER UserIdentity
        The user to modify permissions for
    .PARAMETER AccessRights
        The new access rights
    .PARAMETER SendNotificationToUser
        Whether to send email notification to the user
    .PARAMETER SharingPermissionFlags
        Additional sharing permission flags
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$CalendarPath,
        
        [Parameter(Mandatory = $true)]
        [String]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$AccessRights,

        [Parameter(Mandatory = $true)]
        [Bool]$SendNotificationToUser,
        
        [Parameter(Mandatory = $false)]
        [String]$SharingPermissionFlags = 'None'
    )
    
    Try {
        $SetParams = @{
            Identity = $CalendarPath
            User = $UserIdentity
            AccessRights = $AccessRights
            SendNotificationToUser = $SendNotificationToUser
            ErrorAction = 'Stop'
        }
        
        # Add sharing permission flags if specified and access rights is Editor
        If ($AccessRights -eq 'Editor' -and $SharingPermissionFlags -ne 'None') {
            $SetParams.SharingPermissionFlags = $SharingPermissionFlags
        }
        
        Write-Progress -Activity 'Modifying calendar permissions' -Status 'Progress ->' -PercentComplete 50
        Set-MailboxFolderPermission @SetParams
        Write-Progress -Activity 'Modifying calendar permissions' -Status 'Progress ->' -Completed
        
        Write-ColorEX -Text '[', '✅ SUCCESS', '] Successfully modified ', "$AccessRights", ' permission for ', "$UserIdentity", ' on ', "$MailboxIdentity", ' calendar' -Color White, Green, White, LightGreen, White, LightBlue, White, LightBlue, White -ANSI8
        
        If ($SendNotificationToUser) {
            Write-ColorEX -Text '[', 'ℹ️ INFO', '] Email notification sent to ', "$UserIdentity", ' about calendar sharing changes' -Color White, Cyan, White, LightBlue, White -ANSI8
        }
        
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error modifying calendar permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Throw
    }
}

Function Remove-CalendarPermission {
    <#
    .SYNOPSIS
        Removes calendar permission for a user
    .PARAMETER MailboxIdentity
        The target mailbox identity
    .PARAMETER CalendarPath
        The specific calendar folder path
    .PARAMETER UserIdentity
        The user to remove permissions from
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String]$MailboxIdentity,
        
        [Parameter(Mandatory = $true)]
        [String]$CalendarPath,
        
        [Parameter(Mandatory = $true)]
        [String]$UserIdentity
    )
    
    Try {
        Write-Progress -Activity 'Removing calendar permissions' -Status 'Progress ->' -PercentComplete 50
        Remove-MailboxFolderPermission -Identity $CalendarPath -User $UserIdentity -Confirm:$false -ErrorAction Stop
        Write-Progress -Activity 'Removing calendar permissions' -Status 'Progress ->' -Completed
        
        Write-ColorEX -Text '[', '✅ SUCCESS', '] Successfully removed calendar permissions for ', "$UserIdentity", ' from ', "$MailboxIdentity", ' calendar' -Color White, Green, White, LightBlue, White, LightBlue, White -ANSI8
        
    } Catch {
        Write-ColorEX -Text '[', '❌ ERROR', '] Error removing calendar permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
        Throw
    }
}

Function Show-MainMenu {
    <#
    .SYNOPSIS
        Shows the main permission management menu and handles user selections
    #>
    [CmdletBinding()]
    Param()
    
    $MainMenuOptions = @(
        '📅 View Calendar Permissions',
        '🆕 Add Calendar Permission', 
        '✏️ Modify Calendar Permission',
        '🗑️ Remove Calendar Permission'
    )
    
    Do {
        Reset-OperationState
        $MenuResult = Show-InteractiveMenu -Title "Exchange Online Calendar Permissions Manager" -Options $MainMenuOptions -AllowQuit $true -ShowStatusBar $false
        
        Switch ($MenuResult.Action) {
            'Select' {
                Switch ($MenuResult.Selected) {
                    0 { Invoke-ViewPermissions }
                    1 { Invoke-AddPermission }
                    2 { Invoke-ModifyPermission }
                    3 { Invoke-RemovePermission }
                }
            }
            'Quit' { 
                $QuitConfirm = Show-InteractiveMenu -Title "Confirm Exit" -Options @('Yes, disconnect and exit', 'No, return to main menu') -ShowStatusBar $false
                If ($QuitConfirm.Selected -eq 0) {
                    Return 'Quit'
                }
            }
            'Cancel' {
                $CancelConfirm = Show-InteractiveMenu -Title "Confirm Exit" -Options @('Yes, disconnect and exit', 'No, return to main menu') -ShowStatusBar $false
                If ($CancelConfirm.Selected -eq 0) {
                    Return 'Quit'
                }
            }
        }
    } While ($true)
}

Function Invoke-ViewPermissions {
    <#
    .SYNOPSIS
        Handles viewing calendar permissions with navigation support
    #>
    [CmdletBinding()]
    Param()
    
    Set-OperationState -Action 'View Permissions' -Step 'Select Target Mailbox'
    
    # Get target mailbox
    $MailboxInput = Get-ValidatedInput -Prompt 'Enter mailbox email address or UPN' -PromptTitle 'Enter Target Mailbox' -ValidationType Email
    If ($MailboxInput.Action -eq 'Cancel') {
        Return
    }
    
    $TargetMailbox = $MailboxInput.Value
    Set-OperationState -TargetMailbox $TargetMailbox -Step 'Verifying Mailbox'
    
    # Verify mailbox exists
    $TargetMailboxInfo = Test-MailboxExists -Identity $TargetMailbox
    If ($null -eq $TargetMailboxInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] Mailbox ', "'$TargetMailbox'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Get available calendars
    Set-OperationState -Step 'Retrieving Available Calendars'
    $AvailableCalendars = Get-AvailableCalendars -MailboxIdentity $TargetMailbox

    If ($AvailableCalendars.Count -eq 0) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] No calendars found for mailbox ', "'$TargetMailbox'" -Color White, Red, White, Yellow -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Select calendar
    Set-OperationState -Step 'Select Calendar'
    $SelectedCalendar = Select-Calendar -Calendars $AvailableCalendars -MailboxDisplayName $TargetMailboxInfo.DisplayName

    If ($null -eq $SelectedCalendar) {
        Return
    }

    # Update the calendar path variable
    $CalendarPath = $SelectedCalendar.Identity
    
    Set-OperationState -Step 'Retrieving Permissions'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] Found mailbox: ', "$($TargetMailboxInfo.DisplayName)", ' (', "$($TargetMailboxInfo.RecipientTypeDetails)", ')' -Color White, Green, White, LightGreen, White, LightBlue, White -ANSI8    
    
    # Get and display permissions
    $Permissions = Get-CalendarPermissions -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath

    Clear-Host
    Show-StatusBar
    Write-ColorEX -Text '  ', 'Calendar Permissions for ', "$($TargetMailboxInfo.DisplayName)" -Color None, LightCyan, LightGreen -Style None, Bold, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('═' * (25 + $TargetMailboxInfo.DisplayName.Length)) -Color None, Cyan -ANSI8
    
    If ($Permissions.Count -le 2) {
        Write-ColorEX -Text '  [', 'ℹ️ INFO', '] No custom permissions found for this calendar' -Color White, Cyan, White -ANSI8 -LinesBefore 1
        Write-ColorEX -Text '  Only default permissions (Default and Anonymous) are present.' -Color LightGray -ANSI8
    }

    # Display permissions with individual Write-ColorEX lines using dynamic colors
    Write-ColorEX -Text '  ', 'User', (' ' * 35), 'Access Rights', (' ' * 15), 'Sharing Permissions' -Color None, LightYellow, None, LightYellow, None, LightYellow -Style None, Bold, None, Bold, None, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('─' * 37), '  ', ('─' * 26), '  ', ('─' * 19) -Color None, LightGray, None, LightGray, None, LightGray -ANSI8
    
    ForEach ($Permission in $Permissions) {
        $UserPadding = [Math]::Max(0, 37 - $Permission.User.Length)
        $AccessPadding = [Math]::Max(0, 23 - $Permission.AccessRights.Length)
        
        # Determine color based on access rights level
        $LineColor = Switch -Regex ($Permission.AccessRights) {
            'Owner'                 { 'Red' }      # Highest level - Red
            'PublishingEditor'      { 'Magenta' }  # High level - Magenta
            'Editor'                { 'Yellow' }   # Medium-High level - Yellow
            'PublishingAuthor'      { 'Cyan' }     # Medium level - Cyan
            'Author'                { 'Green' }    # Medium level - Green
            'NonEditingAuthor'      { 'Blue' }     # Medium-Low level - Blue
            'Reviewer'              { 'White' }         # Low level - White
            'Contributor'           { 'LightGray' }     # Low level - Gray
            'AvailabilityOnly'      { 'DarkGray' }      # Minimal - Dark Gray
            'LimitedDetails'        { 'DarkGray' }      # Minimal - Dark Gray
            'Default|Anonymous'     { 'DarkGray' }      # System defaults - Dark Gray
            Default                 { 'LightGray' }     # Unknown - Light Gray
        }
        
        Write-ColorEX -Text '  ', "$($Permission.User)", (' ' * $UserPadding), '  ', "$($Permission.AccessRights)", (' ' * $AccessPadding), '  ', "$($Permission.SharingPermissionFlags)" -Color None, $LineColor, None, None, $LineColor, None, None, $LineColor -Style None, Bold, None, None, Bold, None, None, None -ANSI8
    }
    
    Write-ColorEX '' -LinesBefore 1
    Write-ColorEX -Text 'Press any key to return to main menu...' -Color DarkYellow -ANSI8
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

Function Invoke-AddPermission {
    <#
    .SYNOPSIS
        Handles adding calendar permissions with full navigation support
    #>
    [CmdletBinding()]
    Param()
    
    Set-OperationState -Action 'Add Permission' -Step 'Select Target Mailbox'
    
    # Get target mailbox
    $TargetMailboxInput = Get-ValidatedInput -Prompt 'Enter target mailbox email address or UPN' -PromptTitle 'Enter Target Mailbox' -ValidationType Email
    If ($TargetMailboxInput.Action -eq 'Cancel') {
        Return
    }
    
    $TargetMailbox = $TargetMailboxInput.Value
    Set-OperationState -TargetMailbox $TargetMailbox -Step 'Verifying Target Mailbox'
    
    # Verify target mailbox exists
    $TargetMailboxInfo = Test-MailboxExists -Identity $TargetMailbox
    If ($null -eq $TargetMailboxInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] Target mailbox ', "'$TargetMailbox'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Get available calendars
    Set-OperationState -Step 'Retrieving Available Calendars'
    $AvailableCalendars = Get-AvailableCalendars -MailboxIdentity $TargetMailbox

    If ($AvailableCalendars.Count -eq 0) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] No calendars found for mailbox ', "'$TargetMailbox'" -Color White, Red, White, Yellow -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Select calendar
    Set-OperationState -Step 'Select Calendar'
    $SelectedCalendar = Select-Calendar -Calendars $AvailableCalendars -MailboxDisplayName $TargetMailboxInfo.DisplayName

    If ($null -eq $SelectedCalendar) {
        Return
    }

    # Update the calendar path variable
    $CalendarPath = $SelectedCalendar.Identity

    Set-OperationState -Step 'Retrieving Current Permissions'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] Target mailbox: ', "$($TargetMailboxInfo.DisplayName)", ' (', "$($TargetMailboxInfo.RecipientTypeDetails)", ')' -Color White, Green, White, LightGreen, White, LightBlue, White -ANSI8
    
    # Show current permissions
    $Permissions = Get-CalendarPermissions -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath

    Clear-Host
    Show-StatusBar
    Write-ColorEX -Text '  ', 'Calendar Permissions for ', "$($TargetMailboxInfo.DisplayName)" -Color None, LightCyan, LightGreen -Style None, Bold, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('═' * (25 + $TargetMailboxInfo.DisplayName.Length)) -Color None, Cyan -ANSI8

    If ($Permissions.Count -le 2) {
        Write-ColorEX -Text '  [', 'ℹ️ INFO', '] No custom permissions found for this calendar' -Color White, Cyan, White -ANSI8 -LinesBefore 1
        Write-ColorEX -Text '  Only default permissions (Default and Anonymous) are present.' -Color LightGray -ANSI8
    }

    # Display permissions with individual Write-ColorEX lines using dynamic colors
    Write-ColorEX -Text '  ', 'User', (' ' * 35), 'Access Rights', (' ' * 15), 'Sharing Permissions' -Color None, LightYellow, None, LightYellow, None, LightYellow -Style None, Bold, None, Bold, None, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('─' * 37), '  ', ('─' * 26), '  ', ('─' * 19) -Color None, LightGray, None, LightGray, None, LightGray -ANSI8
    
    ForEach ($Permission in $Permissions) {
        $UserPadding = [Math]::Max(0, 37 - $Permission.User.Length)
        $AccessPadding = [Math]::Max(0, 23 - $Permission.AccessRights.Length)
        
        # Determine color based on access rights level
        $LineColor = Switch -Regex ($Permission.AccessRights) {
            'Owner'                 { 'Red' }      # Highest level - Red
            'PublishingEditor'      { 'Magenta' }  # High level - Magenta
            'Editor'                { 'Yellow' }   # Medium-High level - Yellow
            'PublishingAuthor'      { 'Cyan' }     # Medium level - Cyan
            'Author'                { 'Green' }    # Medium level - Green
            'NonEditingAuthor'      { 'Blue' }     # Medium-Low level - Blue
            'Reviewer'              { 'White' }         # Low level - White
            'Contributor'           { 'LightGray' }     # Low level - Gray
            'AvailabilityOnly'      { 'DarkGray' }      # Minimal - Dark Gray
            'LimitedDetails'        { 'DarkGray' }      # Minimal - Dark Gray
            'Default|Anonymous'     { 'DarkGray' }      # System defaults - Dark Gray
            Default                 { 'LightGray' }     # Unknown - Light Gray
        }
        
        Write-ColorEX -Text '  ', "$($Permission.User)", (' ' * $UserPadding), '  ', "$($Permission.AccessRights)", (' ' * $AccessPadding), '  ', "$($Permission.SharingPermissionFlags)" -Color None, $LineColor, None, None, $LineColor, None, None, $LineColor -Style None, Bold, None, None, Bold, None, None, None -ANSI8
    }
    

    Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8 -LinesBefore 1
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
    Set-OperationState -Step 'Select User for Permissions'
    
    # Get user to grant permissions to
    $UserInput = Get-ValidatedInput -Prompt 'Enter user email address or UPN to grant permissions to' -PromptTitle 'Enter Mailbox Identifier' -ValidationType Email
    If ($UserInput.Action -eq 'Cancel') {
        Return
    }
    
    $UserToGrant = $UserInput.Value
    Set-OperationState -UserMailbox $UserToGrant -Step 'Verifying User'
    
    # Verify user exists
    $UserInfo = Test-MailboxExists -Identity $UserToGrant
    If ($null -eq $UserInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] User ', "'$UserToGrant'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }
    
    Set-OperationState -Step 'Select Access Rights'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] User to grant permissions: ', "$($UserInfo.DisplayName)" -Color White, Green, White, LightGreen -ANSI8
    
    # Select access rights
    $AccessRightsOptions = @(
        'Owner ─ Full control of the calendar',
        'PublishingEditor ─ Create, read, modify, and delete items and folders',
        'Editor ─ Create, read, modify, and delete items',
        'PublishingAuthor ─ Create and read items and folders; modify and delete own items',
        'Author ─ Create and read items; modify and delete own items',
        'NonEditingAuthor ─ Create and read items; delete own items',
        'Reviewer ─ Read items only',
        'Contributor ─ Create items only',
        'AvailabilityOnly ─ View free/busy information only',
        'LimitedDetails ─ View free/busy and subject information'
    )
    
    $AccessMenuResult = Show-InteractiveMenu -Title "Select Access Rights" -Options $AccessRightsOptions -AllowBack $true
    If ($AccessMenuResult.Action -ne 'Select') {
        Return
    }
    
    $AccessRights = $AccessRightsOptions[$AccessMenuResult.Selected].Split(' ─ ')[0]
    Set-OperationState -Step 'Configure Delegate Settings'
    
    # For Editor permissions, ask about delegate permissions
    $SharingFlags = 'None'
    If ($AccessRights -eq 'Editor') {
        $DelegateOptions = @(
            'No ─ Standard Editor permissions only',
            'Yes ─ Make delegate (receive meeting invites)', 
            'Yes ─ Make delegate with private item access'
        )
        
        $DelegateMenuResult = Show-InteractiveMenu -Title "Delegate Permissions" -Options $DelegateOptions -AllowBack $true
        If ($DelegateMenuResult.Action -ne 'Select') {
            Return
        }
        
        Switch ($DelegateMenuResult.Selected) {
            1 { $SharingFlags = 'Delegate' }
            2 { $SharingFlags = 'Delegate,CanViewPrivateItems' }
        }
    }

    # Show sample email and confirm email notification
    Set-OperationState -Step 'Configure Email Notification'
    $NotificationOptions = @(
        'Yes ─ Send email notification to user',
        'No ─ Do not send email notification'
    )

    # Show sample email and confirm email notification
    Set-OperationState -Step 'Configure Email Notification'
    $NotificationOptions = @(
        'Yes ─ Send email notification to user',
        'No ─ Do not send email notification'
    )

    $SampleEmailParams = @{
        FromDisplayName = $TargetMailboxInfo.DisplayName
        FromEmailAddress = $TargetMailbox
        ToDisplayName = $UserInfo.DisplayName
        ToEmailAddress = $UserToGrant
        AccessRights = $AccessRights
        SharingPermissionFlags = $SharingFlags
    }
    
    $NotificationMenuResult = Show-InteractiveMenu -Title "Email Notification" -Options $NotificationOptions -AllowBack $true -ShowSampleEmail $true -SampleEmailParams $SampleEmailParams
    If ($NotificationMenuResult.Action -ne 'Select') {
        Return
    }
    
    $SendNotificationToUser = ($NotificationMenuResult.Selected -eq 0)
    
    # Confirm with summary screen and add permissions
    Set-OperationState -Step 'Confirm Changes'
    $ConfirmOptions = @(
        '✅ Yes ─ Add these permissions',
        '❌ No ─ Return to main menu'
    )

    $SummaryWidth = 70
    $SummaryBorderTop = "┏" + ("━" * ($SummaryWidth - 2)) + "┓"
    $SummaryBorderMiddle = "┣" + ("━" * ($SummaryWidth - 2)) + "┫"
    $SummaryBorderBottom = "┗" + ("━" * ($SummaryWidth - 2)) + "┛"

    # Create summary content as script blocks
    $SummaryContent = @(
        { Write-ColorEX -Text "  $SummaryBorderTop" -Color LightYellow -ANSI8 -LinesBefore 1 },
        { 
            $TitlePadding = $SummaryWidth - 18 - 4
            Write-ColorEX -Text "  ┃ ", 'Permission Summary', (" " * $TitlePadding), " ┃" -Color LightYellow, LightYellow, None, LightYellow -Style None, @('Bold', 'Underline'), None, None -ANSI8
        },
        { Write-ColorEX -Text "  $SummaryBorderMiddle" -Color LightYellow -ANSI8 },
        {
            $TargetPadding = $SummaryWidth - $TargetMailboxInfo.DisplayName.Length - 20
            Write-ColorEX -Text "  ┃ ", 'Target Mailbox: ', "$($TargetMailboxInfo.DisplayName)", (" " * $TargetPadding), " ┃" -Color LightYellow, LightGray, White, None, LightYellow -ANSI8
        },
        {
            $UserPadding = $SummaryWidth - $UserInfo.DisplayName.Length - 10
            Write-ColorEX -Text "  ┃ ", 'User: ', "$($UserInfo.DisplayName)", (" " * $UserPadding), " ┃" -Color LightYellow, LightGray, White, None, LightYellow -ANSI8
        },
        {
            $AccessPadding = $SummaryWidth - $AccessRights.Length - 19
            Write-ColorEX -Text "  ┃ ", 'Access Rights: ', "$AccessRights", (" " * $AccessPadding), " ┃" -Color LightYellow, LightGray, LightGreen, None, LightYellow -ANSI8
        }
    )
    
    If ($SharingFlags -ne 'None') {
        $SummaryContent += {
            $DelegatePadding = $SummaryWidth - $SharingFlags.Length - 22
            Write-ColorEX -Text "  ┃ ", 'Delegate Permissions: ', "$SharingFlags", (" " * $DelegatePadding), " ┃" -Color LightYellow, LightGray, LightBlue, None, LightYellow -ANSI8
        }
    }
    
    # Add notification status
    $NotificationText = If ($SendNotificationToUser) { 'Yes' } Else { 'No' }
    $NotificationColor = If ($SendNotificationToUser) { 'LightGreen' } Else { 'LightRed' }
    $SummaryContent += {
        $NotificationPadding = $SummaryWidth - $NotificationText.Length - 29
        Write-ColorEX -Text "  ┃ ", 'Send Email Notification: ', "$NotificationText", (" " * $NotificationPadding), " ┃" -Color LightYellow, LightGray, $NotificationColor, None, LightYellow -ANSI8
    }
    
    $SummaryContent += { Write-ColorEX -Text "  $SummaryBorderBottom" -Color LightYellow -ANSI8 -LinesAfter 1 }

    $ConfirmResult = Show-InteractiveMenu -Title 'Confirm permission changes' -Options $ConfirmOptions -AllowBack $true -SummaryContent $SummaryContent
    
    If ($ConfirmResult.Selected -eq 0) {
        Try {
            Add-CalendarPermission -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath -UserIdentity $UserToGrant -AccessRights $AccessRights -SendNotificationToUser $SendNotificationToUser -SharingPermissionFlags $SharingFlags
            
            Set-OperationState -Step 'Permissions Changed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX -Text '  ', '✅ Permission Added Successfully!' -Color None, Green -Style None, @('Bold', 'Underline') -ANSI8 -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8 -LinesBefore 1
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } Catch {
            Set-OperationState -Step 'Permissions Failed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX -Text '[', '❌ ERROR', '] Failed to add permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8 -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8 -LinesBefore 1
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
}

Function Invoke-ModifyPermission {
    <#
    .SYNOPSIS
        Handles modifying existing calendar permissions with navigation support
    #>
    [CmdletBinding()]
    Param()
    
    Set-OperationState -Action 'Modify Permission' -Step 'Select Target Mailbox'
    
    # Get target mailbox
    $TargetMailboxInput = Get-ValidatedInput -Prompt 'Enter target mailbox email address or UPN' -PromptTitle 'Enter Target Mailbox' -ValidationType Email
    If ($TargetMailboxInput.Action -eq 'Cancel') {
        Return
    }
    
    $TargetMailbox = $TargetMailboxInput.Value
    Set-OperationState -TargetMailbox $TargetMailbox -Step 'Verifying Target Mailbox'
    
    # Verify target mailbox exists
    $TargetMailboxInfo = Test-MailboxExists -Identity $TargetMailbox
    If ($null -eq $TargetMailboxInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] Target mailbox ', "'$TargetMailbox'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Get available calendars
    Set-OperationState -Step 'Retrieving Available Calendars'
    $AvailableCalendars = Get-AvailableCalendars -MailboxIdentity $TargetMailbox

    If ($AvailableCalendars.Count -eq 0) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] No calendars found for mailbox ', "'$TargetMailbox'" -Color White, Red, White, Yellow -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Select calendar
    Set-OperationState -Step 'Select Calendar'
    $SelectedCalendar = Select-Calendar -Calendars $AvailableCalendars -MailboxDisplayName $TargetMailboxInfo.DisplayName

    If ($null -eq $SelectedCalendar) {
        Return
    }

    # Update the calendar path variable
    $CalendarPath = $SelectedCalendar.Identity
    
    Set-OperationState -Step 'Retrieving Current Permissions'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] Target mailbox: ', "$($TargetMailboxInfo.DisplayName)", ' (', "$($TargetMailboxInfo.RecipientTypeDetails)", ')' -Color White, Green, White, LightGreen, White, LightBlue, White -ANSI8
      
    # Show current permissions
    $Permissions = Get-CalendarPermissions -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath

    Clear-Host
    Show-StatusBar
    Write-ColorEX '' -LinesBefore 1
    Write-ColorEX -Text '  ', 'Calendar Permissions for ', "$($TargetMailboxInfo.DisplayName)" -Color None, LightCyan, LightGreen -Style None, Bold, Bold -ANSI8
    Write-ColorEX -Text '  ', ('═' * (25 + $TargetMailboxInfo.DisplayName.Length)) -Color None, Cyan -ANSI8
    Write-ColorEX ''

    If ($Permissions.Count -le 2) {
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '  [', 'ℹ️ INFO', '] No custom permissions found for this calendar' -Color White, Orange, White -ANSI8
        Write-ColorEX -Text '  Only default permissions (Default and Anonymous) are present.' -Color LightGray -ANSI8
        Write-ColorEX '' -LinesBefore 1
    }

    # Display permissions with individual Write-ColorEX lines using dynamic colors
    Write-ColorEX -Text '  ', 'User', (' ' * 35), 'Access Rights', (' ' * 15), 'Sharing Permissions' -Color None, LightYellow, None, LightYellow, None, LightYellow -Style None, Bold, None, Bold, None, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('─' * 37), '  ', ('─' * 26), '  ', ('─' * 19) -Color None, LightGray, None, LightGray, None, LightGray -ANSI8
    
    ForEach ($Permission in $Permissions) {
        $UserPadding = [Math]::Max(0, 37 - $Permission.User.Length)
        $AccessPadding = [Math]::Max(0, 23 - $Permission.AccessRights.Length)
        
        # Determine color based on access rights level
        $LineColor = Switch -Regex ($Permission.AccessRights) {
            'Owner'                 { 'Red' }      # Highest level - Red
            'PublishingEditor'      { 'Magenta' }  # High level - Magenta
            'Editor'                { 'Yellow' }   # Medium-High level - Yellow
            'PublishingAuthor'      { 'Cyan' }     # Medium level - Cyan
            'Author'                { 'Green' }    # Medium level - Green
            'NonEditingAuthor'      { 'Blue' }     # Medium-Low level - Blue
            'Reviewer'              { 'White' }         # Low level - White
            'Contributor'           { 'LightGray' }     # Low level - Gray
            'AvailabilityOnly'      { 'DarkGray' }      # Minimal - Dark Gray
            'LimitedDetails'        { 'DarkGray' }      # Minimal - Dark Gray
            'Default|Anonymous'     { 'DarkGray' }      # System defaults - Dark Gray
            Default                 { 'LightGray' }     # Unknown - Light Gray
        }
        
        Write-ColorEX -Text '  ', "$($Permission.User)", (' ' * $UserPadding), '  ', "$($Permission.AccessRights)", (' ' * $AccessPadding), '  ', "$($Permission.SharingPermissionFlags)" -Color None, $LineColor, None, None, $LineColor, None, None, $LineColor -Style None, Bold, None, None, Bold, None, None, None -ANSI8
    }
    

    Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8 -LinesBefore 1
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
    Set-OperationState -Step 'Select User to Modify'
    
    # Get user to modify permissions for
    $RemoveMenuResult = Show-InteractiveMenu -Title "Select user to modify calender permissions for $TargetMailbox" -Options ($Permissions).User -AllowBack $true
    If ($RemoveMenuResult.Action -ne 'Select') {
        Invoke-ModifyPermission
    }

    $UserToModify = $Permissions[$RemoveMenuResult.Selected].User

    Set-OperationState -UserMailbox $UserToModify -Step 'Verifying User'
    
    # Verify user exists
    $UserInfo = Test-MailboxExists -Identity $UserToModify
    If ($null -eq $UserInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] User ', "'$UserToModify'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }
    
    Set-OperationState -Step 'Select New Access Rights'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] User to modify permissions: ', "$($UserInfo.DisplayName)" -Color White, Green, White, LightGreen -ANSI8
    
    # Select new access rights (same options as Add)
    $AccessRightsOptions = @(
        'Owner ─ Full control of the calendar',
        'PublishingEditor ─ Create, read, modify, and delete items and folders',
        'Editor ─ Create, read, modify, and delete items',
        'PublishingAuthor ─ Create and read items and folders; modify and delete own items',
        'Author ─ Create and read items; modify and delete own items',
        'NonEditingAuthor ─ Create and read items; delete own items',
        'Reviewer ─ Read items only',
        'Contributor ─ Create items only',
        'AvailabilityOnly ─ View free/busy information only',
        'LimitedDetails ─ View free/busy and subject information'
    )
    
    $AccessMenuResult = Show-InteractiveMenu -Title "Select New Access Rights" -Options $AccessRightsOptions -AllowBack $true
    If ($AccessMenuResult.Action -ne 'Select') {
        Return
    }
    
    $AccessRights = $AccessRightsOptions[$AccessMenuResult.Selected].Split(' ─ ')[0]
    Set-OperationState -Step 'Configure Delegate Settings'
    
    # For Editor permissions, ask about delegate permissions
    $SharingFlags = 'None'
    If ($AccessRights -eq 'Editor') {
        $DelegateOptions = @(
            'No ─ Standard Editor permissions only',
            'Yes ─ Make delegate (receive meeting invites)', 
            'Yes ─ Make delegate with private item access'
        )
        
        $DelegateMenuResult = Show-InteractiveMenu -Title "Delegate Permissions" -Options $DelegateOptions -AllowBack $true
        If ($DelegateMenuResult.Action -ne 'Select') {
            Return
        }
        
        Switch ($DelegateMenuResult.Selected) {
            1 { $SharingFlags = 'Delegate' }
            2 { $SharingFlags = 'Delegate,CanViewPrivateItems' }
        }
    }

    # sample email if user selected to send notification and confirm email notification
    Set-OperationState -Step 'Configure Email Notification'
    $NotificationOptions = @(
        'Yes ─ Send email notification to user',
        'No ─ Do not send email notification'
    )

    # sample email if user selected to send notification and confirm email notification
    Set-OperationState -Step 'Configure Email Notification'
    $NotificationOptions = @(
        'Yes ─ Send email notification to user',
        'No ─ Do not send email notification'
    )

    $SampleEmailParams = @{
        FromDisplayName = $TargetMailboxInfo.DisplayName
        FromEmailAddress = $TargetMailbox
        ToDisplayName = $UserInfo.DisplayName
        ToEmailAddress = $UserToModify
        AccessRights = $AccessRights
        SharingPermissionFlags = $SharingFlags
    }
    
    $NotificationMenuResult = Show-InteractiveMenu -Title "Email Notification" -Options $NotificationOptions -AllowBack $true -ShowSampleEmail $true -SampleEmailParams $SampleEmailParams
    If ($NotificationMenuResult.Action -ne 'Select') {
        Return
    }
    
    $SendNotificationToUser = ($NotificationMenuResult.Selected -eq 0)
    
    # Confirm and modify permission
    Set-OperationState -Step 'Confirm Changes'
    $ConfirmOptions = @(
        '✅ Yes ─ Modify these permissions',
        '❌ No ─ Return to main menu'
    )
    
    $SummaryWidth = 80
    $SummaryBorderTop = "┏" + ("━" * ($SummaryWidth - 2)) + "┓"
    $SummaryBorderMiddle = "┣" + ("━" * ($SummaryWidth - 2)) + "┫"
    $SummaryBorderBottom = "┗" + ("━" * ($SummaryWidth - 2)) + "┛"
    
    # Create summary content as script blocks
    $SummaryContent = @(
        { Write-ColorEX -Text "  $SummaryBorderTop" -Color LightYellow -ANSI8 -LinesBefore 1 },
        { 
            $TitlePadding = $SummaryWidth - 30 - 5
            Write-ColorEX -Text "  ┃ ", 'Permission Modification Summary', (" " * $TitlePadding), " ┃" -Color LightYellow, LightYellow, None, LightYellow -Style None, @('Bold', 'Underline'), None, None -ANSI8
        },
        { Write-ColorEX -Text "  $SummaryBorderMiddle" -Color LightYellow -ANSI8 },
        {
            $TargetPadding = $SummaryWidth - $TargetMailboxInfo.DisplayName.Length - 20
            Write-ColorEX -Text "  ┃ ", 'Target Mailbox: ', "$($TargetMailboxInfo.DisplayName)", (" " * $TargetPadding), " ┃" -Color LightYellow, LightGray, White, None, LightYellow -ANSI8
        },
        {
            $UserPadding = $SummaryWidth - $UserInfo.DisplayName.Length - 10
            Write-ColorEX -Text "  ┃ ", 'User: ', "$($UserInfo.DisplayName)", (" " * $UserPadding), " ┃" -Color LightYellow, LightGray, White, None, LightYellow -ANSI8
        },
        {
            $AccessPadding = $SummaryWidth - $AccessRights.Length - 23
            Write-ColorEX -Text "  ┃ ", 'New Access Rights: ', "$AccessRights", (" " * $AccessPadding), " ┃" -Color LightYellow, LightGray, LightGreen, None, LightYellow -ANSI8
        }
    )
    
    If ($SharingFlags -ne 'None') {
        $SummaryContent += {
            $DelegatePadding = $SummaryWidth - $SharingFlags.Length - 24
            Write-ColorEX -Text "  ┃ ", 'Delegate Permissions: ', "$SharingFlags", (" " * $DelegatePadding), " ┃" -Color LightYellow, LightGray, LightBlue, None, LightYellow -ANSI8
        }
    }
    
    # Add notification status
    $NotificationText = If ($SendNotificationToUser) { 'Yes' } Else { 'No' }
    $NotificationColor = If ($SendNotificationToUser) { 'LightGreen' } Else { 'LightRed' }
    $SummaryContent += {
        $NotificationPadding = $SummaryWidth - $NotificationText.Length - 29
        Write-ColorEX -Text "  ┃ ", 'Send Email Notification: ', "$NotificationText", (" " * $NotificationPadding), " ┃" -Color LightYellow, LightGray, $NotificationColor, None, LightYellow -ANSI8
    }
    
    $SummaryContent += { Write-ColorEX -Text "  $SummaryBorderBottom" -Color LightYellow -ANSI8 -LinesAfter 1 }
    
    $ConfirmResult = Show-InteractiveMenu -Title "Confirm Permission Modification" -Options $ConfirmOptions -ShowStatusBar $true -SummaryContent $SummaryContent

    If ($ConfirmResult.Selected -eq 0) {
        Try {
            Set-CalendarPermission -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath -UserIdentity $UserToModify -AccessRights $AccessRights -SendNotificationToUser $SendNotificationToUser -SharingPermissionFlags $SharingFlags
            
            Set-OperationState -Step 'Permissions Changed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text '  ', '✅ Permission Modified Successfully!' -Color None, Green -Style None, @('Bold', 'Underline') -ANSI8
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } Catch {
            Set-OperationState -Step 'Permissions Failed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text '[', '❌ ERROR', '] Failed to modify permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
}

Function Invoke-RemovePermission {
    <#
    .SYNOPSIS
        Handles removing calendar permissions with navigation support
    #>
    [CmdletBinding()]
    Param()
    
    Set-OperationState -Action 'Remove Permission' -Step 'Select Target Mailbox'
    
    # Get target mailbox
    $TargetMailboxInput = Get-ValidatedInput -Prompt 'Enter target mailbox email address or UPN' -PromptTitle 'Enter Target Mailbox' -ValidationType Email
    If ($TargetMailboxInput.Action -eq 'Cancel') {
        Return
    }
    
    $TargetMailbox = $TargetMailboxInput.Value
    Set-OperationState -TargetMailbox $TargetMailbox -Step 'Verifying Target Mailbox'
    
    # Verify target mailbox exists
    $TargetMailboxInfo = Test-MailboxExists -Identity $TargetMailbox
    If ($null -eq $TargetMailboxInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] Target mailbox ', "'$TargetMailbox'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Get available calendars
    Set-OperationState -Step 'Retrieving Available Calendars'
    $AvailableCalendars = Get-AvailableCalendars -MailboxIdentity $TargetMailbox

    If ($AvailableCalendars.Count -eq 0) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] No calendars found for mailbox ', "'$TargetMailbox'" -Color White, Red, White, Yellow -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }

    # Select calendar
    Set-OperationState -Step 'Select Calendar'
    $SelectedCalendar = Select-Calendar -Calendars $AvailableCalendars -MailboxDisplayName $TargetMailboxInfo.DisplayName

    If ($null -eq $SelectedCalendar) {
        Return
    }

    # Update the calendar path variable
    $CalendarPath = $SelectedCalendar.Identity
    
    Set-OperationState -Step 'Retrieving Current Permissions'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] Target mailbox: ', "$($TargetMailboxInfo.DisplayName)", ' (', "$($TargetMailboxInfo.RecipientTypeDetails)", ')' -Color White, Green, White, LightGreen, White, LightBlue, White -ANSI8
    
    # Show current permissions
    $Permissions = Get-CalendarPermissions -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath

    Clear-Host
    Show-StatusBar
    Write-ColorEX -Text '  ', 'Calendar Permissions for ', "$($TargetMailboxInfo.DisplayName)" -LinesBefore 1
    Write-ColorEX -Text '  ', ('═' * (25 + $TargetMailboxInfo.DisplayName.Length)) -Color None, Cyan -ANSI8

    If ($Permissions.Count -le 2) {
        Write-ColorEX -Text '  [', 'ℹ️ INFO', '] No custom permissions found for this calendar' -Color White, Cyan, White -ANSI8 -LinesBefore 1
        Write-ColorEX -Text '  Only default permissions (Default and Anonymous) are present.' -Color LightGray -ANSI8
    }     

    # Display permissions with individual Write-ColorEX lines using dynamic colors
    Write-ColorEX -Text '  ', 'User', (' ' * 35), 'Access Rights', (' ' * 15), 'Sharing Permissions' -Color None, LightYellow, None, LightYellow, None, LightYellow -Style None, Bold, None, Bold, None, Bold -ANSI8 -LinesBefore 1
    Write-ColorEX -Text '  ', ('─' * 37), '  ', ('─' * 26), '  ', ('─' * 19) -Color None, LightGray, None, LightGray, None, LightGray -ANSI8
    
    ForEach ($Permission in $Permissions) {
        $UserPadding = [Math]::Max(0, 37 - $Permission.User.Length)
        $AccessPadding = [Math]::Max(0, 23 - $Permission.AccessRights.Length)
        
        # Determine color based on access rights level
        $LineColor = Switch -Regex ($Permission.AccessRights) {
            'Owner'                 { 'Red' }      # Highest level - Red
            'PublishingEditor'      { 'Magenta' }  # High level - Magenta
            'Editor'                { 'Yellow' }   # Medium-High level - Yellow
            'PublishingAuthor'      { 'Cyan' }     # Medium level - Cyan
            'Author'                { 'Green' }    # Medium level - Green
            'NonEditingAuthor'      { 'Blue' }     # Medium-Low level - Blue
            'Reviewer'              { 'White' }         # Low level - White
            'Contributor'           { 'LightGray' }     # Low level - Gray
            'AvailabilityOnly'      { 'DarkGray' }      # Minimal - Dark Gray
            'LimitedDetails'        { 'DarkGray' }      # Minimal - Dark Gray
            'Default|Anonymous'     { 'DarkGray' }      # System defaults - Dark Gray
            Default                 { 'LightGray' }     # Unknown - Light Gray
        }
        
        Write-ColorEX -Text '  ', "$($Permission.User)", (' ' * $UserPadding), '  ', "$($Permission.AccessRights)", (' ' * $AccessPadding), '  ', "$($Permission.SharingPermissionFlags)" -Color None, $LineColor, None, None, $LineColor, None, None, $LineColor -Style None, Bold, None, None, Bold, None, None, None -ANSI8
    }


    Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8 -LinesBefore 1
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
    Set-OperationState -Step 'Select User to Remove'
    
    # Get user to remove permissions from
    $RemoveMenuResult = Show-InteractiveMenu -Title "Select user to remove calender permissions from $TargetMailbox" -Options ($Permissions | Where-Object {$_.User -ne 'Default' -and $_.User -ne 'Anonymous'}).User -AllowBack $true
    If ($RemoveMenuResult.Action -ne 'Select') {
        Invoke-RemovePermission
    }

    $UserToRemove = $Permissions[$RemoveMenuResult.Selected + 2].User

    Set-OperationState -UserMailbox $UserToRemove -Step 'Verifying User'
    
    # Verify user exists
    $UserInfo = Test-MailboxExists -Identity $UserToRemove
    If ($null -eq $UserInfo) {
        Clear-Host
        Show-StatusBar
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text '[', '❌ ERROR', '] User ', "'$UserToRemove'", ' not found in Exchange Online' -Color White, Red, White, Yellow, White -ANSI8
        Write-ColorEX '' -LinesBefore 1
        Write-ColorEX -Text 'Press any key to continue...' -Color Yellow -ANSI8
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Return
    }
    
    Set-OperationState -Step 'Confirm Removal'
    Write-ColorEX -Text '[', '✅ SUCCESS', '] User to remove permissions: ', "$($UserInfo.DisplayName)" -Color White, Green, White, LightGreen -ANSI8
    
    # Confirm removal with summary and warning
    $ConfirmOptions = @(
        '⚠️ Yes ─ Remove ALL permissions for this user',
        '❌ No ─ Return to main menu'
    )
    
    $SummaryWidth = 70
    $SummaryBorderTop = "┏" + ("━" * ($SummaryWidth - 2)) + "┓"
    $SummaryBorderMiddle = "┣" + ("━" * ($SummaryWidth - 2)) + "┫"
    $SummaryBorderBottom = "┗" + ("━" * ($SummaryWidth - 2)) + "┛"
    
    # Confirm removal with summary and warning
    $ConfirmOptions = @(
        '⚠️ Yes ─ Remove ALL permissions for this user',
        '❌ No ─ Return to main menu'
    )
    
    $SummaryWidth = 70
    $SummaryBorderTop = "┏" + ("━" * ($SummaryWidth - 2)) + "┓"
    $SummaryBorderMiddle = "┣" + ("━" * ($SummaryWidth - 2)) + "┫"
    $SummaryBorderBottom = "┗" + ("━" * ($SummaryWidth - 2)) + "┛"
    
    # Create summary content as script blocks
    $SummaryContent = @(
        { Write-ColorEX -Text "  $SummaryBorderTop" -Color LightRed -ANSI8 -LinesBefore 1 },
        { 
            $TitlePadding = $SummaryWidth - 25 - 5
            Write-ColorEX -Text "  ┃ ", 'Permission Removal Summary', (" " * $TitlePadding), " ┃" -Color LightRed, LightRed, None, LightRed -Style None, @('Bold', 'Underline'), None, None -ANSI8
        },
        { Write-ColorEX -Text "  $SummaryBorderMiddle" -Color LightRed -ANSI8 },
        {
            $TargetPadding = $SummaryWidth - $TargetMailboxInfo.DisplayName.Length - 20
            Write-ColorEX -Text "  ┃ ", 'Target Mailbox: ', "$($TargetMailboxInfo.DisplayName)", (" " * $TargetPadding), " ┃" -Color LightRed, LightGray, White, None, LightRed -ANSI8
        },
        {
            $UserPadding = $SummaryWidth - $UserInfo.DisplayName.Length - 20
            Write-ColorEX -Text "  ┃ ", 'User to Remove: ', "$($UserInfo.DisplayName)", (" " * $UserPadding), " ┃" -Color LightRed, LightGray, White, None, LightRed -ANSI8
        },
        { Write-ColorEX -Text "  $SummaryBorderMiddle" -Color LightRed -ANSI8 },
        {
            $WarningPadding = $SummaryWidth - 50 - 8
            Write-ColorEX -Text "  ┃ ", '⚠️ WARNING: This will remove ', 'ALL', ' calendar permissions!', (" " * $WarningPadding), " ┃" -Color LightRed, Orange, Red, Orange, None, LightRed -Style None, None, Bold, None, None, None -ANSI8
        },
        {
            $NoticePadding = $SummaryWidth - 52 - 9
            Write-ColorEX -Text "  ┃ ", 'The user will no longer have any access to this calendar.', (" " * $NoticePadding), " ┃" -Color LightRed, LightGray, None, LightRed -ANSI8
        },
        { Write-ColorEX -Text "  $SummaryBorderBottom" -Color LightRed -ANSI8 -LinesAfter 1 }
    )
    
    $ConfirmResult = Show-InteractiveMenu -Title "Confirm Permission Removal" -Options $ConfirmOptions -ShowStatusBar $true -SummaryContent $SummaryContent

    If ($ConfirmResult.Selected -eq 0) {
        Try {
            Remove-CalendarPermission -MailboxIdentity $TargetMailbox -CalendarPath $CalendarPath -UserIdentity $UserToRemove
            
            Set-OperationState -Step 'Permissions Changed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text '  ', '✅ Permission Removed Successfully!' -Color None, Green -Style None, @('Bold', 'Underline') -ANSI8
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } Catch {
            Set-OperationState -Step 'Permissions Failed'
            Clear-Host
            Show-StatusBar
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text '[', '❌ ERROR', '] Failed to remove permission: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
            Write-ColorEX '' -LinesBefore 1
            Write-ColorEX -Text 'Press any key to return to main menu...' -Color Yellow -ANSI8
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
}

Function Write-Banner {
    Clear-Host
    Write-ColorEX -Text "┏","━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━","┓" -Color Blue, Blue, Blue -BackgroundColor Black, Black, Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","                                            .-.                                           ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","                                            ─#─              #.    ─+                     ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","    ....           .       ...      ...     ─#─  .          =#:..          ...      ..    ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","   +===*#─  ",".:","     #*  *#++==*#:   +===**:  ─#─ .#*    ─#─ =*#+++. +#.  ─*+==+*. .*+─=*.  ","┃" -Color Blue,Blue,Cyan,Blue,Blue -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","    .::.+#  ",".:","     #*  *#    .#+   .::..**  ─#─  .#+  ─#=   =#:    +#. =#:       :#+:     ","┃" -Color Blue,Blue,Cyan,Blue,Blue -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","  =#=──=##. ",".:","     #*  *#     #+  **───=##  ─#─   .#+─#=    =#:    +#. **          :=**.  ","┃" -Color Blue,Blue,Cyan,Blue,Blue -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","  **.  .*#. ",".:.","   =#=  *#     #+ :#=   :##  ─#─    :##=     ─#─    +#. :#*::  .:  ::  .#= ","┃" -Color Blue,Blue,Cyan,Blue,Blue -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","  ─+++──=      .==:   ==     =─  .=++= ─==  :=:    .#=       ─++=  ─=    :=+++─  :=++= ─  ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","                                                  .#+                                     ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┃","                                                  *+                                      ","┃" -Color Blue,Blue,Blue -BackgroundColor Black,Black,Black -HorizontalCenter -ANSI8
    Write-ColorEX -Text "┗","━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━","┛" -Color Blue, Blue, Blue -BackgroundColor Black, Black, Black -HorizontalCenter -ANSI8
    Write-ColorEX '' -LinesBefore 1
    Write-ColorEX -Text '┏', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '┓' -Color Blue, Blue, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┃', '             Exchange Online Calendar Manager            ', '┃' -Color Blue, White, Blue -Style None, Bold, None -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┃', '                 Version 1.1 - 2025─07─16                ', '┃' -Color Blue, LightGray, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┃', '                   Author: Mark Newton                   ', '┃' -Color Blue, Gray, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┗', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '┛' -Color Blue, Blue, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX '' -LinesBefore 1
}

Function Initialize-Script {
    <#
    .SYNOPSIS
        Initializes the script by checking requirements and establishing connections
    #>
    
    # Display banner
    Write-Banner
    
    # Set initial window title
    Update-WindowTitle
    
    # Check PowerShell version
    If ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-ColorEX -Text '[', '❌ ERROR', '] PowerShell 5.1 or later is required. Current version: ', "$($PSVersionTable.PSVersion)" -Color White, Red, White, Yellow -ANSI8
        Throw "Unsupported PowerShell version"
    }
    
    Write-ColorEX -Text '[', '✅ SUCCESS', '] PowerShell version check passed: ', "$($PSVersionTable.PSVersion)" -Color White, Green, White, LightGreen -ANSI8
    
    # Check and install Exchange Online module if needed
    If (-not (Test-ModuleInstalled)) {
        Write-ColorEX -Text '[', '⚠️ WARNING', '] Exchange Online Management module not found or outdated' -Color White, Orange, White -ANSI8
        Install-ExchangeModule
    } Else {
        Import-Module -Name $Script:ModuleName -Force -ErrorAction Stop
        Start-Sleep 2
    }
    
    # Get admin UPN if not provided
    If ([String]::IsNullOrWhiteSpace($AdminUPN)) {
        $AdminUPNInput = Get-ValidatedInput -Prompt 'Enter Exchange Admin UPN for connecting to Exchange Online' -PromptTitle 'Exchange Admin UPN' -ValidationType Email
        If ($AdminUPNInput.Action -eq 'Cancel') {
            Write-ColorEX -Text '[', 'ℹ️ INFO', '] Connection cancelled by user' -Color White, Cyan, White -ANSI8
            Exit 0
        }
        $AdminUPN = $AdminUPNInput.Value
    }
    
    # Connect to Exchange Online
    Connect-ExchangeOnlineService -UserPrincipalName $AdminUPN
    
    Write-ColorEX -Text '[', '✅ SUCCESS', '] Script initialization completed successfully' -Color White, Green, White -ANSI8
    Write-ColorEX '' -LinesBefore 1
    Write-ColorEX -Text 'Press any key to continue to main menu...' -Color Yellow -ANSI8
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

Function Show-GoodbyeMessage {
    <#
    .SYNOPSIS
        Shows a professional goodbye message when exiting
    #>
    
    # Display banner
    Write-Banner

    If ($Script:IsConnected) {
        Write-ColorEX -Text '[', 'ℹ️ INFO', '] Disconnecting from Exchange Online...' -Color White, Cyan, White -ANSI8 -HorizontalCenter
        Try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-ColorEX -Text '[', '✅ SUCCESS', '] Disconnected successfully' -Color White, Green, White -ANSI8 -HorizontalCenter
        } Catch {
            Write-ColorEX -Text '[', '⚠️ WARNING', '] Disconnect may have failed, but session will timeout automatically' -Color White, Orange, White -ANSI8 -HorizontalCenter
        }
    }

    Write-ColorEX '' -LinesBefore 2
    Write-ColorEX -Text '┏', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '┓' -Color Blue, Blue, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┃', '                        Session Complete                      ', '┃' -Color Blue, White, Blue -Style None, Bold, None -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┃', '                  Stay Classy, A','u','nalytics 🥂                  ', '┃' -Color Blue, Blue, Cyan, Blue, Blue -ANSI8 -HorizontalCenter
    Write-ColorEX -Text '┗', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '┛' -Color Blue, Blue, Blue -ANSI8 -HorizontalCenter
    
    # Reset window title
    $Host.UI.RawUI.WindowTitle = 'PowerShell'

    Write-ColorEX -Text 'Press any key to exit...' -Color Red -ANSI8
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

#endregion

# ================================
# ===           MAIN           ===
# ================================
#region Main Execution

Try {
    Initialize-Script
    
    Do {
        $MainMenuResult = Show-MainMenu
        If ($MainMenuResult -eq 'Quit') {
            Break
        }
    } While ($true)
    
    Show-GoodbyeMessage
    
} Catch {
    Write-ColorEX -Text '[', 'FATAL ERROR', '] Script execution failed: ', "$($_.Exception.Message)" -Color White, Red, White, LightRed -ANSI8
    Write-ColorEX '' -LinesBefore 1
    Write-ColorEX -Text 'Press any key to exit...' -Color Red -ANSI8
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit 1
} Finally {
    # Ensure cleanup
    If ($Script:IsConnected) {
        Try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        } Catch {
            # Ignore disconnection errors in cleanup
        }
    }
    
    # Reset window title
    $Host.UI.RawUI.WindowTitle = 'PowerShell'
}

#endregion
