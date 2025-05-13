# Termination Script: Active Directory Group Membership Verification
# 
# Input: Workday Termination Report CSV file (Employee_ID, Termination_Date)
# 
# This PowerShell script is used to verify Active Directory (AD) group memberships and account statuses 
# for employees listed in a CSV file, typically after a termination process. 
# 
# Key Features:
# - Allows the user to select a CSV file containing employee data (Employee_ID and Termination_Date) via a file dialog or manual input.
# - Converts Excel files (.xlsx) to CSV format if necessary.
# - Ensures the script is run with administrative privileges and installs the RSAT tools if required.
# - Imports the Active Directory module and validates its availability.
# - Searches AD for users matching the Employee_IDs in the CSV file and retrieves their group memberships and account statuses.
# - Filters users based on membership in specific groups ("SMS Users" and "Sugarbush-SUG-RTP") or active account status.
# - Outputs the results, including termination dates, in a table format.
# - Provides an option to re-run the script or exit.
param (
    [string]$csvFilePath
)

function getCSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select Termination Report"
    }
    $null = $FileBrowser.ShowDialog()
    $FileName = $FileBrowser.FileName

    if ($FileName -match '\.xlsx$') {
        $Excel = New-Object -ComObject Excel.Application
        $wb = $Excel.Workbooks.Open($FileName)
        $FileName = $FileName.Replace(".xlsx", ".csv")
        foreach ($ws in $wb.Worksheets) {
            $ws.SaveAs($FileName, 6)
        }
        $Excel.Quit()
    }

    # Remove junk lines if needed
    $firstline = Get-Content $FileName -First 6
    if ($firstline -match "Terminations") {
        (Get-Content $FileName | Select-Object -Skip 6) | Set-Content $FileName
    }

    return $FileName
}

function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

do {
    Clear-Variable -Name csvFilePath, csvData, employeeIDs, terminationDates, results -ErrorAction SilentlyContinue

    Write-Host "Step 1: Installing RSAT tools (if not already installed)..."
    if (Test-IsAdmin) {
        try {
            Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
            Write-Host "RSAT tools installed successfully."
        } catch {
            Write-Warning "Could not install RSAT tools: $_"
        }
    } else {
        Write-Warning "You are not running as Administrator. RSAT install skipped."
    }

    Write-Host "Step 2: Importing Active Directory module..."
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Host "Error: Active Directory module is not available." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    }

    Write-Host "Step 3: Selecting CSV file..."
    $useExplorer = Read-Host "Would you like to open file through file explorer? (Y/N)"
    $csvFilePath = if ($useExplorer -eq "Y") { getCSVFile } else { Read-Host "Enter full path to CSV" }
    $csvFilePath = $csvFilePath -replace '[\"'']', ''

    if (-not (Test-Path $csvFilePath)) {
        Write-Host "Error: File not found." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit
    }

    try {
        $csvData = Import-Csv $csvFilePath
    } catch {
        Write-Host "Error: Failed to import CSV." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit
    }

    if (-not ($csvData | Get-Member -Name Employee_ID)) {
        Write-Host "Missing 'Employee_ID' column." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit
    }

    if (-not ($csvData | Get-Member -Name Termination_Date)) {
        Write-Warning "Missing 'Termination_Date' column. Will show N/A in output."
    }

    $employeeIDs = $csvData | Select-Object -ExpandProperty Employee_ID
    $terminationDates = $csvData | Select-Object -ExpandProperty Termination_Date
    $results = @()

    Write-Host "Searching Active Directory..."

    for ($i = 0; $i -lt $employeeIDs.Count; $i++) {
        $employeeID = $employeeIDs[$i]
        $terminationDate = if ($terminationDates[$i]) { $terminationDates[$i] } else { "N/A" }

        $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties SamAccountName, Name, Enabled, MemberOf, Title

        if ($user) {
            $userGroups = $user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
            $isSMS = $userGroups -contains "SMS Users"
            $isRTP = $userGroups -contains "Sugarbush-SUG-RTP"

            $jobTitle = $user.Title
            $jobPrefix = if ($jobTitle -match "^(supervisor|manager|director)") { $matches[1].ToLower() } else { $jobTitle.ToLower() }

            $action = if ($jobPrefix -in @("supervisor", "manager", "director")) {
                "Suspend with Delegation"
            } else {
                "Suspend with no Delegate"
            }

            if ($isSMS -or $isRTP -or $user.Enabled) {
                $results += [PSCustomObject]@{
                    Username            = $user.SamAccountName
                    EmployeeID          = $employeeID
                    "Termination Date"  = $terminationDate
                    "Job Title"         = $jobTitle
                    "SMS Users"         = if ($isSMS) { "Yes" } else { "No" }
                    "Sugarbush-SUG-RTP" = if ($isRTP) { "Yes" } else { "No" }
                    "Active Account"    = if ($user.Enabled) { "Yes" } else { "No" }
                    Action              = $action
                }
            }
        } else {
            Write-Host "User not found for Employee ID: $employeeID" -ForegroundColor Yellow
        }
    }

    Write-Host "`nResults:"
    $results | Format-Table -AutoSize

    $choice = Read-Host "`nPress R to re-run or Enter to exit"
} while ($choice -eq "R")
