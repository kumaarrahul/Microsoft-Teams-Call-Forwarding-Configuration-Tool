<#
.SYNOPSIS
    Microsoft Teams Call Forwarding Configuration Tool with Voicemail Backup
    
.DESCRIPTION
    This script provides an interactive menu-driven tool to configure call forwarding settings
    in Microsoft Teams. It supports both immediate and unanswered call forwarding options and
    creates timestamped backups of all settings including voicemail configuration before making changes.
    
.PARAMETER None
    This script does not accept parameters.
    
.INPUTS
    The script requires a CSV file named "users.csv" in the same directory with the following columns:
    - Email: The email address of the Teams user
    - ForwardingNumber: The phone number where calls will be forwarded
    
.OUTPUTS
    - backup_[timestamp].csv: Backup file containing original settings before changes
    - current_settings_[timestamp].csv: File containing updated settings after changes
    - scriptlog_[timestamp].log: Detailed log of all operations performed
    
.NOTES
    Version:        1.1
    Author:         Rahul Kumaar
    Creation Date:  March 24, 2025
    Updated:        Enhanced with voicemail settings backup
    
    Requirements:
    - Microsoft Teams PowerShell module must be installed
    - User must have appropriate permissions to modify Teams call settings
    - Run with administrative privileges in PowerShell
    
.EXAMPLE
    .\TeamsCallForwarding.ps1
    
    Runs the script and displays the interactive menu.
#>
# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get the current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get current date and time for file naming
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Define file paths relative to the script directory
$csvFilePath = Join-Path -Path $scriptDir -ChildPath "users.csv"
$backupFilePath = Join-Path -Path $scriptDir -ChildPath "backup_$timestamp.csv"
$exportFilePath = Join-Path -Path $scriptDir -ChildPath "current_settings_$timestamp.csv"
$logFilePath = Join-Path -Path $scriptDir -ChildPath "scriptlog_$timestamp.log"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFilePath -Value "$timestamp - $message"
}

# Function to display menu and get user choice
function Show-Menu {
    Clear-Host
    Write-Host "================ TEAMS CALL FORWARDING CONFIGURATION ================"
    Write-Host "1: Configure Immediate Transfer"
    Write-Host "2: Configure Unanswered Transfer"
    Write-Host "3: Backup Current Settings Only"
    Write-Host "4: Exit"
    Write-Host "=================================================================="
    
    $choice = Read-Host "Enter your choice (1-4)"
    return $choice
}

# Function to backup current settings including voicemail
function Backup-Settings {
    Log-Message "Starting backup of current settings including voicemail."
    Write-Host "Starting backup operation..." -ForegroundColor Cyan
    
    $backup = @()
    
    if (Test-Path -Path $csvFilePath) {
        $users = Import-Csv -Path $csvFilePath
        $totalUsers = $users.Count
        $currentUser = 0
        
        Write-Host "Found $totalUsers users to process" -ForegroundColor Cyan
        
        foreach ($user in $users) {
            $currentUser++
            Write-Host "[$currentUser/$totalUsers] Processing backup for: $($user.Email)" -ForegroundColor White
            Log-Message "Backing up settings for user: $($user.Email)"
            
            # Use the Get-CurrentSettings function to retrieve all settings
            $backupResult = Get-CurrentSettings -userEmail $user.Email
            
            if ($backupResult -ne $null) {
                $backup += $backupResult
                Log-Message "Successfully backed up all settings for: $($user.Email)"
                Write-Host "    Backup completed for: $($user.Email)" -ForegroundColor Green
            } else {
                Log-Message "Failed to backup settings for user: $($user.Email)"
                Write-Host "    Failed to backup: $($user.Email)" -ForegroundColor Red
            }
            
            Start-Sleep -Seconds 1
        }
        
        Write-Host "Exporting backup data to file..." -ForegroundColor Cyan
        # Export the backup settings to a CSV file
        $backup | Export-Csv -Path $backupFilePath -NoTypeInformation
        Log-Message "Backup settings (including voicemail) exported to $backupFilePath."
        Write-Host "Backup operation completed successfully!" -ForegroundColor Green
        
        return $backup
    } else {
        $errorMessage = "Input file not found: $csvFilePath"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Message $errorMessage
        return $null
    }
}

# Function to get current settings including voicemail
function Get-CurrentSettings {
    param (
        [string]$userEmail
    )
    
    try {
        Write-Host "  Processing user: $userEmail" -ForegroundColor Yellow
        
        # Get call forwarding settings
        Write-Host "    Retrieving call forwarding settings..." -ForegroundColor Gray
        $currentSettings = Get-CsUserCallingSettings -Identity $userEmail
        
        # Get voicemail settings
        Write-Host "    Retrieving voicemail settings..." -ForegroundColor Gray
        $voicemailSettings = $null
        try {
            $voicemailSettings = Get-CsOnlineVoicemailUserSettings -Identity $userEmail
        } catch {
            Log-Message "Warning: Could not retrieve voicemail settings for: $userEmail - $_"
        }
        
        Write-Host "    Settings retrieved successfully" -ForegroundColor Green
        
        $result = [PSCustomObject]@{
            Email = $userEmail
            # Call Forwarding Settings
            ImmediateEnabled = $currentSettings.IsForwardingEnabled
            ImmediateTargetType = $currentSettings.ForwardingTargetType
            ImmediateTarget = $currentSettings.ForwardingTarget
            UnansweredEnabled = $currentSettings.IsUnansweredEnabled
            UnansweredTargetType = $currentSettings.UnansweredTargetType
            UnansweredTarget = $currentSettings.UnansweredTarget
            UnansweredDelay = $currentSettings.UnansweredDelay
            # Voicemail Settings
            VoicemailEnabled = if ($voicemailSettings) { $voicemailSettings.VoicemailEnabled } else { "N/A" }
            CallAnsweringRulesEnabled = if ($voicemailSettings) { $voicemailSettings.CallAnsweringRulesEnabled } else { "N/A" }
            DefaultGreetingPromptOverwrite = if ($voicemailSettings) { $voicemailSettings.DefaultGreetingPromptOverwrite } else { "N/A" }
            DefaultOofGreetingPromptOverwrite = if ($voicemailSettings) { $voicemailSettings.DefaultOofGreetingPromptOverwrite } else { "N/A" }
            OofGreetingEnabled = if ($voicemailSettings) { $voicemailSettings.OofGreetingEnabled } else { "N/A" }
            OofGreetingFollowAutomaticRepliesEnabled = if ($voicemailSettings) { $voicemailSettings.OofGreetingFollowAutomaticRepliesEnabled } else { "N/A" }
            OofGreetingFollowCalendarEnabled = if ($voicemailSettings) { $voicemailSettings.OofGreetingFollowCalendarEnabled } else { "N/A" }
            PromptLanguage = if ($voicemailSettings) { $voicemailSettings.PromptLanguage } else { "N/A" }
            ShareData = if ($voicemailSettings) { $voicemailSettings.ShareData } else { "N/A" }
            TransferTarget = if ($voicemailSettings) { $voicemailSettings.TransferTarget } else { "N/A" }
            VoicemailTranscriptionEnabled = if ($voicemailSettings) { $voicemailSettings.VoicemailTranscriptionEnabled } else { "N/A" }
            VoicemailTranscriptionProfanityMaskingEnabled = if ($voicemailSettings) { $voicemailSettings.VoicemailTranscriptionProfanityMaskingEnabled } else { "N/A" }
            VoicemailTranscriptionTranslationEnabled = if ($voicemailSettings) { $voicemailSettings.VoicemailTranscriptionTranslationEnabled } else { "N/A" }
        }
        
        return $result
    } catch {
        Log-Message "Error retrieving current settings for user: $userEmail - $_"
        return $null
    }
}

# Function to configure immediate transfer
function Configure-ImmediateTransfer {
    Log-Message "Starting configuration of immediate transfer."
    Write-Host "Starting immediate transfer configuration..." -ForegroundColor Cyan
    
    if (Test-Path -Path $csvFilePath) {
        $backup = Backup-Settings
        
        if ($backup -ne $null) {
            $results = @()
            
            $users = Import-Csv -Path $csvFilePath
            $totalUsers = $users.Count
            $currentUser = 0
            
            Write-Host "Configuring immediate transfer for $totalUsers users..." -ForegroundColor Cyan
            
            foreach ($user in $users) {
                try {
                    $currentUser++
                    Write-Host "[$currentUser/$totalUsers] Configuring immediate transfer for: $($user.Email)" -ForegroundColor White
                    Log-Message "Configuring immediate transfer for user: $($user.Email)"
                    
                    Write-Host "    Applying call forwarding settings..." -ForegroundColor Yellow
                    Set-CsUserCallingSettings -Identity $user.Email `
                                             -IsForwardingEnabled $true `
                                             -ForwardingType Immediate `
                                             -ForwardingTargetType SingleTarget `
                                             -ForwardingTarget $user.ForwardingNumber
                    
                    Write-Host "    Retrieving updated settings..." -ForegroundColor Yellow
                    # Get updated settings including voicemail
                    $updatedSettings = Get-CurrentSettings -userEmail $user.Email
                    $updatedSettings.ForwardingNumber = $user.ForwardingNumber
                    
                    $results += $updatedSettings
                    Log-Message "Successfully configured immediate transfer for: $($user.Email)"
                    Write-Host "    Configuration completed for: $($user.Email)" -ForegroundColor Green
                    
                    Start-Sleep -Seconds 2
                } catch {
                    Log-Message "Error configuring immediate transfer for user: $($user.Email) - $_"
                    Write-Host "    Error configuring: $($user.Email) - $_" -ForegroundColor Red
                }
            }
            
            Write-Host "Exporting updated settings to file..." -ForegroundColor Cyan
            # Export the updated settings to a CSV file
            $results | Export-Csv -Path $exportFilePath -NoTypeInformation
            Log-Message "Updated settings (including voicemail) exported to $exportFilePath."
            
            Write-Host "Immediate transfer configuration completed." -ForegroundColor Green
            Log-Message "Immediate transfer configuration completed."
        }
    } else {
        $errorMessage = "Input file not found: $csvFilePath"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Message $errorMessage
    }
}

# Function to configure unanswered transfer
function Configure-UnansweredTransfer {
    Log-Message "Starting configuration of unanswered transfer."
    
    if (Test-Path -Path $csvFilePath) {
        $backup = Backup-Settings
        
        if ($backup -ne $null) {
            $results = @()
            
            $users = Import-Csv -Path $csvFilePath
            
            foreach ($user in $users) {
                try {
                    Log-Message "Configuring unanswered transfer for user: $($user.Email)"
                    
                    Set-CsUserCallingSettings -Identity $user.Email `
                                             -IsUnansweredEnabled $true `
                                             -UnansweredTargetType SingleTarget `
                                             -UnansweredTarget $user.ForwardingNumber `
                                             -UnansweredDelay "00:00:10"
                    
                    # Get updated settings including voicemail
                    $updatedSettings = Get-CurrentSettings -userEmail $user.Email
                    $updatedSettings.ForwardingNumber = $user.ForwardingNumber
                    
                    $results += $updatedSettings
                    Log-Message "Successfully configured unanswered transfer for: $($user.Email)"
                    
                    Start-Sleep -Seconds 2
                } catch {
                    Log-Message "Error configuring unanswered transfer for user: $($user.Email) - $_"
                }
            }
            
            # Export the updated settings to a CSV file
            $results | Export-Csv -Path $exportFilePath -NoTypeInformation
            Log-Message "Updated settings (including voicemail) exported to $exportFilePath."
            
            Write-Host "Unanswered transfer configuration completed." -ForegroundColor Green
            Log-Message "Unanswered transfer configuration completed."
        }
    } else {
        $errorMessage = "Input file not found: $csvFilePath"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Message $errorMessage
    }
}

# Display script information
Write-Host "Teams Call Forwarding Configuration Tool with Voicemail Backup" -ForegroundColor Cyan
Write-Host "--------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Input File: $csvFilePath" -ForegroundColor Yellow
Write-Host "Backup File: $backupFilePath" -ForegroundColor Yellow
Write-Host "Export File: $exportFilePath" -ForegroundColor Yellow
Write-Host "Log File: $logFilePath" -ForegroundColor Yellow
Write-Host ""

# Main program
Log-Message "Script started with voicemail backup enhancement."

do {
    $choice = Show-Menu
    
    switch ($choice) {
        "1" {
            Configure-ImmediateTransfer
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "2" {
            Configure-UnansweredTransfer
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "3" {
            $result = Backup-Settings
            if ($result -ne $null) {
                Write-Host "Backup completed. Settings (including voicemail) saved to $backupFilePath" -ForegroundColor Green
            }
            Log-Message "Backup only operation completed."
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "4" {
            Write-Host "Exiting script." -ForegroundColor Cyan
            Log-Message "Script exited by user."
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
} while ($choice -ne "4")

# Log the end of the script
Log-Message "Script completed."