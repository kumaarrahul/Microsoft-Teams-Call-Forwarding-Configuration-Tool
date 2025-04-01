# Teams Call Forwarding Configuration Tool

A PowerShell utility for managing Microsoft Teams call forwarding settings for multiple users simultaneously. This tool offers both immediate and unanswered call forwarding configuration with automatic backups and detailed logging.

## Features

- **Menu-driven interface** for easy navigation and operation
- **Immediate call forwarding** - configure calls to be forwarded immediately
- **Unanswered call forwarding** - configure calls to be forwarded after a timeout (10 seconds)
- **Automatic backups** with timestamps of all settings before making changes
- **Detailed logging** of all operations with timestamps
- **Portable design** - input/output files are located in the same directory as the script

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Teams PowerShell module installed
- Appropriate administrative permissions to modify Teams call settings
- CSV file with user information in the required format

## Installation

1. Download the `TeamsCallForwarding.ps1` script
2. Create a CSV file called `users.csv` in the same directory (see format below)
3. Run the script in PowerShell with administrative privileges

```powershell
.\TeamsCallForwarding.ps1
```

## Input File Format

The script requires a CSV file named `users.csv` in the same directory with the following columns:

- `Email`: The email address of the Teams user
- `ForwardingNumber`: The phone number where calls will be forwarded

Example:
```csv
Email,ForwardingNumber
user1@company.com,+15551234567
user2@company.com,+15559876543
```

The `ForwardingNumber` should be in E.164 format (with a "+" prefix followed by country code and phone number).

## Generated Files

The script generates the following files in the same directory, each with a timestamp for identification:

- **backup_[timestamp].csv**: Original settings before changes were made
- **current_settings_[timestamp].csv**: Updated settings after changes were applied
- **scriptlog_[timestamp].log**: Detailed log of all operations performed

## Usage

1. Run the script in PowerShell
2. Select an option from the menu:
   - Option 1: Configure Immediate Transfer
   - Option 2: Configure Unanswered Transfer
   - Option 3: Backup Current Settings Only
   - Option 4: Exit

## Example

```powershell
PS C:\Scripts> .\TeamsCallForwarding.ps1

Teams Call Forwarding Configuration Tool
----------------------------------------
Input File: C:\Scripts\users.csv
Backup File: C:\Scripts\backup_20250401_143022.csv
Export File: C:\Scripts\current_settings_20250401_143022.csv
Log File: C:\Scripts\script_20250401_143022.log

================ TEAMS CALL FORWARDING CONFIGURATION ================
1: Configure Immediate Transfer
2: Configure Unanswered Transfer
3: Backup Current Settings Only
4: Exit
==================================================================
```

## Troubleshooting

- Ensure you have the Teams PowerShell module installed: `Install-Module -Name MicrosoftTeams`
- Verify that the users.csv file exists and has the correct format
- Check the log file for detailed error messages
- Ensure you have administrative permissions in Microsoft Teams

## License

[MIT License](LICENSE)

## Contributors

- Rahul Kumaar

