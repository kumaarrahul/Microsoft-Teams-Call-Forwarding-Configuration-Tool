<#
.SYNOPSIS
    Microsoft Teams Call Forwarding Configuration Tool
    
.DESCRIPTION
    This script provides an interactive menu-driven tool to configure call forwarding settings
    in Microsoft Teams. It supports both immediate and unanswered call forwarding options and
    creates timestamped backups of all settings before making changes.
    
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
    Version:        1.0
    Author:         Rahul Kumaar
    Creation Date:  March 24, 2025
    
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

# Function to backup current settings
function Backup-Settings {
    Log-Message "Starting backup of current settings."
    
    $backup = @()
    
    if (Test-Path -Path $csvFilePath) {
        $users = Import-Csv -Path $csvFilePath
        
        foreach ($user in $users) {
            try {
                Log-Message "Backing up settings for user: $($user.Email)"
                
                $currentSettings = Get-CsUserCallingSettings -Identity $user.Email
                
                $backupResult = [PSCustomObject]@{
                    Email = $user.Email
                    ImmediateEnabled = $currentSettings.IsForwardingEnabled
                    ImmediateTargetType = $currentSettings.ForwardingTargetType
                    ImmediateTarget = $currentSettings.ForwardingTarget
                    UnansweredEnabled = $currentSettings.IsUnansweredEnabled
                    UnansweredTargetType = $currentSettings.UnansweredTargetType
                    UnansweredTarget = $currentSettings.UnansweredTarget
                    UnansweredDelay = $currentSettings.UnansweredDelay
                }
                
                $backup += $backupResult
                Log-Message "Successfully backed up settings for: $($user.Email)"
                
                Start-Sleep -Seconds 1
            } catch {
                Log-Message "Error backing up settings for user: $($user.Email) - $_"
            }
        }
        
        # Export the backup settings to a CSV file
        $backup | Export-Csv -Path $backupFilePath -NoTypeInformation
        Log-Message "Backup settings exported to $backupFilePath."
        
        return $backup
    } else {
        $errorMessage = "Input file not found: $csvFilePath"
        Write-Host $errorMessage -ForegroundColor Red
        Log-Message $errorMessage
        return $null
    }
}

# Function to configure immediate transfer
function Configure-ImmediateTransfer {
    Log-Message "Starting configuration of immediate transfer."
    
    if (Test-Path -Path $csvFilePath) {
        $backup = Backup-Settings
        
        if ($backup -ne $null) {
            $results = @()
            
            $users = Import-Csv -Path $csvFilePath
            
            foreach ($user in $users) {
                try {
                    Log-Message "Configuring immediate transfer for user: $($user.Email)"
                    
                    Set-CsUserCallingSettings -Identity $user.Email `
                                             -IsForwardingEnabled $true `
                                             -ForwardingType Immediate `
                                             -ForwardingTargetType SingleTarget `
                                             -ForwardingTarget $user.ForwardingNumber
                    
                    $updatedSettings = Get-CsUserCallingSettings -Identity $user.Email
                    
                    $result = [PSCustomObject]@{
                        Email = $user.Email
                        ForwardingNumber = $user.ForwardingNumber
                        ImmediateEnabled = $updatedSettings.IsForwardingEnabled
                        ImmediateTargetType = $updatedSettings.ForwardingTargetType
                        ImmediateTarget = $updatedSettings.ForwardingTarget
                        UnansweredEnabled = $updatedSettings.IsUnansweredEnabled
                        UnansweredTargetType = $updatedSettings.UnansweredTargetType
                        UnansweredTarget = $updatedSettings.UnansweredTarget
                        UnansweredDelay = $updatedSettings.UnansweredDelay
                    }
                    
                    $results += $result
                    Log-Message "Successfully configured immediate transfer for: $($user.Email)"
                    
                    Start-Sleep -Seconds 2
                } catch {
                    Log-Message "Error configuring immediate transfer for user: $($user.Email) - $_"
                }
            }
            
            # Export the updated settings to a CSV file
            $results | Export-Csv -Path $exportFilePath -NoTypeInformation
            Log-Message "Updated settings exported to $exportFilePath."
            
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
                    
                    $updatedSettings = Get-CsUserCallingSettings -Identity $user.Email
                    
                    $result = [PSCustomObject]@{
                        Email = $user.Email
                        ForwardingNumber = $user.ForwardingNumber
                        ImmediateEnabled = $updatedSettings.IsForwardingEnabled
                        ImmediateTargetType = $updatedSettings.ForwardingTargetType
                        ImmediateTarget = $updatedSettings.ForwardingTarget
                        UnansweredEnabled = $updatedSettings.IsUnansweredEnabled
                        UnansweredTargetType = $updatedSettings.UnansweredTargetType
                        UnansweredTarget = $updatedSettings.UnansweredTarget
                        UnansweredDelay = $updatedSettings.UnansweredDelay
                    }
                    
                    $results += $result
                    Log-Message "Successfully configured unanswered transfer for: $($user.Email)"
                    
                    Start-Sleep -Seconds 2
                } catch {
                    Log-Message "Error configuring unanswered transfer for user: $($user.Email) - $_"
                }
            }
            
            # Export the updated settings to a CSV file
            $results | Export-Csv -Path $exportFilePath -NoTypeInformation
            Log-Message "Updated settings exported to $exportFilePath."
            
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
Write-Host "Teams Call Forwarding Configuration Tool" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host "Input File: $csvFilePath" -ForegroundColor Yellow
Write-Host "Backup File: $backupFilePath" -ForegroundColor Yellow
Write-Host "Export File: $exportFilePath" -ForegroundColor Yellow
Write-Host "Log File: $logFilePath" -ForegroundColor Yellow
Write-Host ""

# Main program
Log-Message "Script started."

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
                Write-Host "Backup completed. Settings saved to $backupFilePath" -ForegroundColor Green
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