# Script Title: Active Directory Manager-Employee Update Script
# 
# Input: Internal CSV file containing Employee ID and Manager details, with columns "Employee ID" and "Manager".
#
# Description:
# This PowerShell script is designed to verify and update the manager-employee relationships in Active Directory (AD) 
# based on data provided in a CSV file. It checks if the current manager assigned to an employee in AD matches the 
# expected manager listed in the CSV file. The script generates a report of discrepancies and provides options to:
# 
# - Print all users with mismatched managers.
# - Update the `Manager` property in AD for users with mismatched managers.
# 
# Key Features:
# - Prompts the user to relaunch the script with administrator privileges if not already running as admin.
# - Allows the user to select a CSV file containing employee and manager data via a file dialog or manual input.
# - Converts Excel files (.xlsx) to CSV format if necessary.
# - Validates the input file for required columns: "Employee ID" and "Manager".
# - Queries AD for users and compares their current manager with the expected manager.
# - Outputs a detailed report of manager-employee relationships, including mismatches.
# - Provides an option to update the `Manager` property in AD for users with mismatched managers.
param (
    [string]$csvFilePath
)

# Prompt user to optionally run as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    $response = Read-Host "This script is not running as Administrator. Would you like to relaunch it with admin rights? Needed to update Manager. (Y/N)"
    if ($response -match '^(Y|y)$') {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
        exit
    } else {
        Write-Host "Continuing without Administrator privileges..." -ForegroundColor Yellow
    }
}

function getCSVFile {
    # Load the required assembly for Windows Forms
    Add-Type -AssemblyName System.Windows.Forms

    # File Selection window
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select the CSV file"
    }
    $null = $FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}

# Function to convert .xlsx to .csv
function ConvertToCSV {
    param (
        [string]$filePath
    )

    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.Workbooks.Open($filePath)
    $tempCsvFilePath = [System.IO.Path]::GetTempFileName().Replace(".tmp", ".csv")
    foreach ($ws in $wb.Worksheets) {
        $ws.SaveAs($tempCsvFilePath, 6) # Save as CSV format
    }
    $Excel.Quit()

    Write-Host "The file has been converted to a temporary CSV file." -ForegroundColor Green
    return $tempCsvFilePath
}

do {
    # Clear all variables to reset the script
    Clear-Variable -Name csvFilePath -ErrorAction SilentlyContinue
    Clear-Variable -Name csvData -ErrorAction SilentlyContinue
    Clear-Variable -Name employeeIDs -ErrorAction SilentlyContinue
    Clear-Variable -Name results -ErrorAction SilentlyContinue

    # Check if script is running as Administrator
    function Test-IsAdmin {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    Write-Host "Step 1: Installing RSAT tools (if not already installed)..."

    if (Test-IsAdmin) {
        try {
            Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
            Write-Host "RSAT tools installed successfully."
        } catch {
            Write-Warning "Could not install RSAT tools: $_"
        }
    } else {
        Write-Warning "You are not running as Administrator. RSAT install skipped. If AD module is missing, rerun this script as Admin."
    }

    # Import the Active Directory module
    Write-Host "Step 2: Importing Active Directory module..."
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Host "Error: Active Directory module is not available. Ensure RSAT is installed." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    }

    # Ask the user if they want to use file explorer
    $openwindowbool = Read-Host "Would you like to open the file through file explorer? (Y/N)"
    if ($openwindowbool -eq "Y") {
        $csvFilePath = getCSVFile
    } else {
        # Prompt for CSV if not passed as argument (drag-and-drop or manual input)
        $csvFilePath = Read-Host "Enter the full path to the CSV file (drag-and-drop or type manually)"
    }

    # Extract and sanitize the actual file path
    if ($csvFilePath -match "([a-zA-Z]:\\[^\`"']+|\\\\[^\\]+\\[^\\]+[^\`"']*)") {
        $csvFilePath = $matches[1] -replace "[`'`"]", ""
    }

    # Convert .xlsx to .csv if needed
    if ($csvFilePath -match '\.xlsx$') {
        $csvFilePath = ConvertToCSV -filePath $csvFilePath
    }

    # Validate path
    if (-Not (Test-Path -Path $csvFilePath)) {
        Write-Host "Error: The specified file does not exist. Please check the path and try again." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    }

    try {
        $csvData = Import-Csv -Path $csvFilePath

        # Normalize column headers
        $normalizedHeaders = $csvData[0].PSObject.Properties.Name | ForEach-Object { $_.Trim() }
        if (-Not ($normalizedHeaders -contains "Employee ID") -or -Not ($normalizedHeaders -contains "Manager")) {
            Write-Host "Error: The CSV file does not contain the required 'Employee ID' or 'Manager' columns." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }
    } catch {
        Write-Host "Error: Unable to import the CSV file. Ensure it is in the correct format." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    }

    $employeeIDs = @($csvData | Select-Object -ExpandProperty "Employee ID")
    $managers = @($csvData | Select-Object -ExpandProperty "Manager")

    $results = @()

    # Querying Active Directory for users
    Write-Host "Step 3: Querying Active Directory for users..."
    for ($i = 0; $i -lt $employeeIDs.Count; $i++) {
        $employeeID = $employeeIDs[$i]
        $managerRaw = $managers[$i]

        # Parse Manager Name and Employee ID
        if ($managerRaw -match "^(.*?)\s*\((\d+)\)$") {
            $managerName = $matches[1].Trim() # Extract name and trim any extra spaces
            $managerID = $matches[2].Trim()   # Extract Employee ID
        } else {
            $managerName = "Invalid Format"
            $managerID = "N/A"
        }

        $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, Manager

        if ($user) {
            # Get the current manager's name and Employee ID if the Manager property is populated
            if ($user.Manager) {
                $currentManager = Get-ADUser -Identity $user.Manager -Properties Name, EmployeeID
                $currentManagerName = $currentManager.Name
                $currentManagerID = $currentManager.EmployeeID
            } else {
                $currentManagerName = "No Manager Assigned"
                $currentManagerID = "N/A"
            }

            # Compare Current Manager ID and New Manager ID
            $managerMatch = if ($currentManagerID -eq $managerID) { "Yes" } else { "No" }

            # Add user details to results
            $results += [PSCustomObject]@{
                Name                 = $user.Name
                Username             = $user.SamAccountName
                EmployeeID           = $user.EmployeeID
                CurrentManager       = $currentManagerName
                CurrentManagerID     = $currentManagerID
                ExpectedManager      = $managerName
                ExpectedManagerID    = $managerID
                ManagerMatch         = $managerMatch
            }
        } else {
            Write-Host "No user found for Employee ID: $employeeID" -ForegroundColor Yellow
        }
    }

    # Display results
    Write-Host "`nResults:"
    $results | Format-Table -AutoSize

    # Option to print non-matching managers
    $printChoice = Read-Host "`nWould you like to print all users with ManagerMatch = 'No'? (Y/N)"
    if ($printChoice -eq "Y") {
        Write-Host "`nUsers with non-matching managers:"
        $results | Where-Object { $_.ManagerMatch -eq "No" } | Format-Table -AutoSize
    }

    # Option to update managers in AD
    $updateChoice = Read-Host "`nWould you like to update managers in AD for users with ManagerMatch = 'No'? (Y/N)"
    if ($updateChoice -eq "Y") {
        $results | Where-Object { $_.ManagerMatch -eq "No" } | ForEach-Object {
            Write-Host "Updating manager for user: $($_.Username)..."
            try {
                Set-ADUser -Identity $_.Username -Manager (Get-ADUser -Filter { EmployeeID -eq $_.ExpectedManagerID }).DistinguishedName
                Write-Host "Manager updated successfully for $($_.Username)." -ForegroundColor Green
            } catch {
                Write-Host "Failed to update manager for $($_.Username): $_" -ForegroundColor Red
            }
        }
    }
    # Prompt to re-run or exit
    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")