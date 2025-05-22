# Script Title: Active Directory Ninjio Training Status Check for All Locations
# 
# Input: Internal Ninjio Report CSV file (Employee ID, Completion Status, Manager)
#
# Description:
# This PowerShell script is designed to verify the Ninjio training completion status of employees across multiple locations 
# and cross-check their account statuses in Active Directory (AD). It allows the user to filter employees by location and 
# generates a detailed report of employees who have not completed their training, along with their account status, manager details, 
# and training completion status.
# 
# Key Features:
# - Prompts the user to select a CSV file containing employee training data via a file dialog or manual input.
# - Converts Excel files (.xlsx) to CSV format if necessary.
# - Validates the input file and skips the first two rows if needed.
# - Allows the user to filter employees by location.
# - Queries AD for users based on Employee ID and retrieves their account status (Enabled/Disabled), manager details, and training completion status.
# - Checks the "Completion Status" column to determine if training is completed ("Completed") or not (anything else).
# - Generates a report of employees for the selected location, including their training status.
# - Displays the report in the console for quick review.
# - Provides an option to re-run the script for another location or exit.
param (
    [string]$csvFilePath
)

# Function to open file explorer and select a file
function getCSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select the CSV or Excel file"
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
    $ws = $wb.Worksheets.Item(1)
    $ws.SaveAs($tempCsvFilePath, 6) # Save as CSV format
    $wb.Close($false)
    $Excel.Quit()
    Write-Host "The file has been converted to a temporary CSV file." -ForegroundColor Green
    return $tempCsvFilePath
}

do {
    # Prompt for file selection
    if (-not $csvFilePath) {
        $csvFilePath = Read-Host "Would you like to open the file through file explorer? (Y/N)"
        if ($csvFilePath -eq "Y") {
            $csvFilePath = getCSVFile
        } else {
            $csvFilePath = Read-Host "Enter the full path to the CSV or Excel file (drag-and-drop or type manually)"
        }
    }

    # Extract and sanitize the actual file path
    if ($csvFilePath -match "([a-zA-Z]:\\[^\`"']+|\\\\[^\\]+\\[^\\]+[^\`"']*)") {
        $csvFilePath = $matches[1] -replace "[`'`"]", ""
    }

    # Validate and process the file
    if ($csvFilePath -match '\.xlsx$') {
        Write-Host "Converting .xlsx file to .csv format..." -ForegroundColor Yellow
        $csvFilePath = ConvertToCSV -filePath $csvFilePath
    }

    if (-not (Test-Path $csvFilePath)) {
        Write-Host "Error: File not found at $csvFilePath" -ForegroundColor Red
        exit
    }

    # Import the CSV file
    try {
        $rawData = Get-Content -Path $csvFilePath
        $filteredData = $rawData | Select-Object -Skip 2 # Skip the first two rows
        $csvData = $filteredData | ConvertFrom-Csv
    } catch {
        Write-Host "Error: Unable to process the file. Ensure it is properly formatted." -ForegroundColor Red
        exit
    }

    # Extract unique locations
    $locations = $csvData | Select-Object -ExpandProperty "Employee's Company" -Unique | Sort-Object
    Write-Host "`nAvailable Locations:"
    for ($i = 0; $i -lt $locations.Count; $i++) {
        Write-Host "$($i + 1). $($locations[$i])"
    }

    # Prompt user to select a location
    $locationIndex = Read-Host "`nEnter the number corresponding to the location you want to check"
    if (-not ($locationIndex -as [int]) -or $locationIndex -lt 1 -or $locationIndex -gt $locations.Count) {
        Write-Host "Invalid selection. Exiting..." -ForegroundColor Red
        exit
    }
    $selectedLocation = $locations[$locationIndex - 1]
    Write-Host "`nYou selected: $selectedLocation" -ForegroundColor Cyan

    # Filter employees by the selected location
    $filteredEmployees = $csvData | Where-Object { $_."Employee's Company" -eq $selectedLocation }

    # Query Active Directory for each employee
    Write-Host "`nChecking Ninjio training status for employees at $selectedLocation..."
    $results = @()
    foreach ($employee in $filteredEmployees) {
    $employeeID = $employee.'Employee ID'
    $completionStatus = $employee.'Completion Status'
    $trainingStatus = if ($completionStatus -eq "Completed") { "Completed" } else { "Not Completed" }

    $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, Name, Enabled, Manager

    if ($user) {
        $managerName = if ($user.Manager) {
            (Get-ADUser -Identity $user.Manager -Properties Name).Name
        } else {
            "No Manager Assigned"
        }

        $results += [PSCustomObject]@{
            Manager                = $managerName
            "Staff to Complete Training" = $user.Name
            EmployeeID             = $user.EmployeeID
            AccountStatus          = if ($user.Enabled) { "Active" } else { "Inactive" }
            TrainingStatus         = $trainingStatus
        }
    } else {
        Write-Host "No user found for Employee ID: $employeeID" -ForegroundColor Yellow
    }
}

# Display results
if ($results.Count -gt 0) {
    Write-Host "`nResults for ${selectedLocation}:"
    $results | Format-Table Manager, 'Staff to Complete Training', EmployeeID, AccountStatus, TrainingStatus -AutoSize

    # Prompt to export results
    $exportChoice = Read-Host "Would you like to export these results to a CSV file in your Downloads folder? (Y/N)"
    if ($exportChoice -eq "Y") {
        $downloadsPath = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportFile = Join-Path $downloadsPath "Ninjio_Training_Results_${($selectedLocation -replace '[^a-zA-Z0-9]', '_')}_$timestamp.csv"
        $results | Export-Csv -Path $exportFile -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $exportFile" -ForegroundColor Green
    }
} else {
    Write-Host "No employees found at ${selectedLocation}." -ForegroundColor Yellow
}

$choice = Read-Host "Press R to re-run the script or Enter to exit"
} while ($choice -eq "R")