# Microsoft Teams Call Forwarding Configuration Tool

A comprehensive PowerShell utility for managing Microsoft Teams call forwarding settings for multiple users simultaneously with enhanced voicemail settings backup functionality. This tool offers both immediate and unanswered call forwarding configuration with automatic backups and detailed logging.

## 🚀 Features

- **Menu-driven interface** for easy navigation and operation
- **Immediate call forwarding** - configure calls to be forwarded immediately to specified numbers
- **Unanswered call forwarding** - configure calls to be forwarded after a timeout period (10 seconds)
- **Comprehensive backup system** with timestamps of all call forwarding AND voicemail settings before making changes
- **Enhanced voicemail settings backup** including transcription, language, greeting settings, and more
- **Detailed logging** of all operations with timestamps for audit trails
- **Portable design** - all input/output files are located in the same directory as the script
- **Error handling** with graceful degradation for voicemail settings that may not be accessible

## 📋 Prerequisites

- **Windows PowerShell 5.1** or later
- **Microsoft Teams PowerShell module** installed and configured
- **Appropriate administrative permissions** to modify Teams call settings and access voicemail configurations
- **CSV file** with user information in the required format

## 📦 Installation

1. **Download the script**: Save `TeamsCallForwarding.ps1` to your desired directory
2. **Install Teams PowerShell module** (if not already installed):
   ```powershell
   Install-Module -Name MicrosoftTeams -Force -AllowClobber
   ```
3. **Create input CSV file**: Create a file named `users.csv` in the same directory as the script
4. **Run with administrative privileges**: Execute the script in PowerShell with administrator permissions

## 📄 CSV File Format

The script requires a CSV file named `users.csv` in the same directory with the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| `Email` | The email address of the Teams user | user1@contoso.com |
| `ForwardingNumber` | The phone number where calls will be forwarded | +15551234567 |

### Example CSV Content:
```csv
Email,ForwardingNumber
user1@contoso.com,+15551234567
user2@contoso.com,+15559876543
user3@contoso.com,+15551112222
```

**Important**: The `ForwardingNumber` should be in E.164 format (with "+" prefix followed by country code and phone number).

## 🔧 Usage

1. **Launch the script**:
   ```powershell
   .\TeamsCallForwarding.ps1
   ```

2. **Connect to Teams**: The script will automatically connect to Microsoft Teams (you may be prompted for authentication)

3. **Select from the menu options**:
   - **Option 1**: Configure Immediate Transfer
   - **Option 2**: Configure Unanswered Transfer  
   - **Option 3**: Backup Current Settings Only
   - **Option 4**: Exit

### Example Session:
```powershell
PS C:\Scripts> .\TeamsCallForwarding.ps1

Teams Call Forwarding Configuration Tool with Voicemail Backup
--------------------------------------------------------------
Input File: C:\Scripts\users.csv
Backup File: C:\Scripts\backup_20250324_143022.csv
Export File: C:\Scripts\current_settings_20250324_143022.csv
Log File: C:\Scripts\scriptlog_20250324_143022.log

================ TEAMS CALL FORWARDING CONFIGURATION ================
1: Configure Immediate Transfer
2: Configure Unanswered Transfer
3: Backup Current Settings Only
4: Exit
==================================================================
Enter your choice (1-4):
```

## 📊 Output Files

The script generates several timestamped files in the same directory for tracking and audit purposes:

### Backup Files
- **`backup_[timestamp].csv`**: Complete backup of original call forwarding AND voicemail settings before any changes
- **`current_settings_[timestamp].csv`**: Updated settings after changes are applied (includes voicemail settings)

### Log Files  
- **`scriptlog_[timestamp].log`**: Detailed log of all operations, errors, and status messages

### Enhanced Backup Content

The backup now includes comprehensive voicemail settings:
- **Call Forwarding Settings**: Immediate/Unanswered forwarding configuration, targets, delays
- **Voicemail Core Settings**: Enabled status, call answering rules
- **Greeting Management**: Default greetings, out-of-office greetings, custom prompts
- **Transcription Features**: Transcription enabled/disabled, profanity masking, translation
- **Language & Localization**: Prompt language settings
- **Advanced Features**: Share data settings, transfer targets, calendar integration

## ⚙️ Configuration Options

### Immediate Transfer (Option 1)
- Forwards all incoming calls immediately to the specified number
- No ringing on the original Teams client
- Ideal for permanent call redirection scenarios

### Unanswered Transfer (Option 2)  
- Allows calls to ring on Teams first
- Forwards to specified number after 10 seconds if unanswered
- Maintains Teams presence while providing backup coverage

### Backup Only (Option 3)
- Creates comprehensive backup without making any changes
- Useful for auditing current configurations
- Includes both call forwarding and voicemail settings

## 🛠️ Troubleshooting

### Common Issues and Solutions

**Teams PowerShell Module Not Found**
```powershell
Install-Module -Name MicrosoftTeams -Force -AllowClobber
```

**Authentication Issues**
- Ensure you have appropriate Teams administrative permissions
- Try disconnecting and reconnecting: `Disconnect-MicrosoftTeams` then `Connect-MicrosoftTeams`

**CSV File Not Found**
- Verify `users.csv` exists in the same directory as the script
- Check file name spelling and format

**Voicemail Settings Warnings**
- Some users may not have voicemail configured - this is normal
- The script will log warnings but continue processing other users
- Voicemail settings will show "N/A" for users without voicemail access

**Permission Errors**
- Run PowerShell as Administrator
- Verify Teams admin permissions for the account being used

### Checking Log Files
Always check the generated log file (`scriptlog_[timestamp].log`) for detailed error messages and operation status.

## 🔒 Security Considerations

- **Backup Strategy**: Always creates backups before making changes
- **Audit Trail**: Comprehensive logging of all operations
- **Minimal Permissions**: Only requests necessary Teams permissions
- **Data Protection**: All files remain local to the script directory

## 📈 Version History

- **v1.1** (March 2025): Enhanced with comprehensive voicemail settings backup and improved error handling
- **v1.0**: Initial release with basic call forwarding configuration

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## 📧 Author

**Rahul Kumaar**  
- GitHub: [@kumaarrahul](https://github.com/kumaarrahul)

## 📄 License

This project is available for use under standard open source practices. Please ensure compliance with your organization's policies when using this tool.

---

*For additional support or feature requests, please create an issue in the GitHub repository.*