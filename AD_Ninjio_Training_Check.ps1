# Script Title: Active Directory Ninjio Training Check Script
#
# Input: Internal Ninjio Report CSV file (Employee ID, Completion Status, Manager)
#
# Description:
# This PowerShell script is designed to verify the training completion status of employees listed in a CSV file 
# and cross-check their account statuses in Active Directory (AD). It generates a report of employees who have 
# not completed their training and provides details such as their account status and supervisor.
# 
# Key Features:
# - Prompts the user to select a CSV file containing employee training data via a file dialog or manual input.
# - Converts Excel files (.xlsx) to CSV format if necessary.
# - Validates the input file and skips the first two rows if needed.
# - Queries AD for users based on Employee ID and retrieves their account status (Enabled/Disabled).
# - Generates a report of employees who have not completed their training.
# - Provides an option to export the report to a CSV file.
# - Displays the report in the console for quick review.
param (
    [string]$csvFilePath
)

# Function to open file explorer and select a file
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

   # Display a loading message
    Write-Host "Processing the file. Please wait..." -ForegroundColor Yellow

    # Convert .xlsx to .csv if needed
    if ($csvFilePath -match '\.xlsx$') {
        Write-Host "Converting .xlsx file to .csv format..." -ForegroundColor Yellow
        $csvFilePath = ConvertToCSV -filePath $csvFilePath
    } else {
        Write-Host "Loading the .csv file..." -ForegroundColor Yellow
    }

    # Validate path
    if (-Not (Test-Path -Path $csvFilePath)) {
        Write-Host "Error: The specified file does not exist. Please check the path and try again." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    } else {
        Write-Host "File path is valid: $csvFilePath" -ForegroundColor Green
    }

    # Import the CSV file, skipping the first two rows
    try {
        Write-Host "Importing the CSV file and skipping the first two rows..." -ForegroundColor Yellow
        $rawData = Get-Content -Path $csvFilePath
        $filteredData = $rawData | Select-Object -Skip 2 # Skip the first two rows
        $csvData = $filteredData | ConvertFrom-Csv
        Write-Host "CSV file imported successfully, starting from the third row." -ForegroundColor Green
    } catch {
        Write-Host "Error: Unable to process the CSV file. Ensure the file is properly formatted." -ForegroundColor Red
        exit
    }

    # Function to generate a unique file path if the file already exists
    function Get-UniqueFilePath {
        param (
            [string]$baseFilePath
        )

        $directory = [System.IO.Path]::GetDirectoryName($baseFilePath)
        $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($baseFilePath)
        $extension = [System.IO.Path]::GetExtension($baseFilePath)
        $counter = 1

        $uniqueFilePath = $baseFilePath
        while (Test-Path $uniqueFilePath) {
            $uniqueFilePath = [System.IO.Path]::Combine($directory, "$fileNameWithoutExtension($counter)$extension")
            $counter++
        }

        return $uniqueFilePath
    }

    # Define the base output file path in the user's Downloads folder
    $baseOutputFilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath("UserProfile"), "Downloads", "Ninjio_Training_Report.csv")

    # Initialize an array to store the report data
    $reportData = @()

    Write-Host "Checking account statuses in Active Directory:" -ForegroundColor Yellow
    foreach ($row in $csvData) {
        $employeeID = $row."Employee ID"
        $supervisor = $row."Manager" # Assuming the column is named "Manager"
        $completionStatus = $row."Completion Status" # Assuming the column is named "Completion Status"

        # Only process rows where Completion Status is empty or not "Completed"
        if ([string]::IsNullOrWhiteSpace($completionStatus) -or $completionStatus -ne "Completed") {
            try {
                # Search for the user in AD using the Employee ID
                $adUser = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties Enabled, Name

                # Determine account status
                $accountStatus = if ($adUser.Enabled -eq $true) { "Enabled" } else { "Disabled" }

                # Add the row to the report data
                $rowData = [PSCustomObject]@{
                    "ACCT STATUS" = $accountStatus
                    "SUPERVISOR" = $supervisor
                    "Staff to complete Training" = $adUser.Name
                }
                $reportData += $rowData

                # Display the row in the console
                $rowData | Format-Table -AutoSize
            } catch {
                # Handle case where user is not found in AD
                $rowData = [PSCustomObject]@{
                    "ACCT STATUS" = "Not Found"
                    "SUPERVISOR" = $supervisor
                    "Staff to complete Training" = "Not Found"
                }
                $reportData += $rowData

                # Display the row in the console
                $rowData | Format-Table -AutoSize
            }
        }
    }

    # Ask the user if they want to export the report
    $exportChoice = Read-Host "Press Y to export the report to a CSV file or any other key to skip"
    if ($exportChoice -eq "Y") {
        # Generate a unique file path if the file already exists
        $outputFilePath = Get-UniqueFilePath -baseFilePath $baseOutputFilePath

        # Export the report data to a CSV file
        $reportData | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8
        Write-Host "Report has been exported to $outputFilePath" -ForegroundColor Green
    } else {
        Write-Host "Report was not exported." -ForegroundColor Yellow
    }

    # Prompt the user to press a button to exit
    Read-Host "Press Enter to exit"

    $choice = Read-Host "Press R to re-run the script or Enter to exit"
} while ($choice -eq "R")