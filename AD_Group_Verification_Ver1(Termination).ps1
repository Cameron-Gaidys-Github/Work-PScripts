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
    # Load the required assembly for Windows Forms
    Add-Type -AssemblyName System.Windows.Forms

    # File Selection window  
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select Job Change Report"
    }
    $null = $FileBrowser.ShowDialog()
    $FileName = $FileBrowser.FileName

    # Convert xlsx to csv if needed
    if ($FileBrowser.FileName -match '.xlsx') {
        $Excel = New-Object -ComObject Excel.Application
        $wb = $Excel.Workbooks.Open($FileBrowser.FileName)
        $FileName = $FileBrowser.FileName.Replace(".xlsx", ".csv")
        foreach ($ws in $wb.Worksheets) {
            $ws.SaveAs($FileName, 6)
        }
        $Excel.Quit()
    }

    # Optional: Remove junk lines if needed
    $firstline = Get-Content $FileName -First 6
    if ($firstline -match "Terminations") {
        (Get-Content $FileName | Select-Object -Skip 6) | Set-Content $FileName
    }

    return $FileName
}

do {
    # Clear all variables to reset the script
    Clear-Variable -Name csvFilePath -ErrorAction SilentlyContinue
    Clear-Variable -Name csvData -ErrorAction SilentlyContinue
    Clear-Variable -Name employeeIDs -ErrorAction SilentlyContinue
    Clear-Variable -Name terminationDates -ErrorAction SilentlyContinue
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

    # Verify the module is imported
    Write-Host "Step 3: Verifying module import..."
    if (Get-Module -Name ActiveDirectory) {
        Write-Host "Step 4: Specify the CSV file to import..."

        # Ask the user if they want to use file explorer
        $openwindowbool = Read-Host "Would you like to open file through file explorer? (Y/N)"
        if ($openwindowbool -eq "Y") {
            $csvFilePath = getCSVFile
        } else {
            # Prompt for CSV if not passed as argument (drag-and-drop or manual input)
            $csvFilePath = Read-Host "Enter the full path to the CSV file (drag-and-drop or type manually)"
        }

        # Remove quotes from the file path
        $csvFilePath = $csvFilePath -replace '[\"'']', ''

        # Validate path
        if (-Not (Test-Path -Path $csvFilePath)) {
            Write-Host "Error: The specified file does not exist. Please check the path and try again." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        try {
            $csvData = Import-Csv -Path $csvFilePath
        } catch {
            Write-Host "Error: Unable to import the CSV file. Ensure it is in the correct format." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        # Check for required columns: Employee_ID and Termination_Date
        if (-Not ($csvData | Get-Member -Name Employee_ID)) {
            Write-Host "Error: The CSV file does not contain the required Employee_ID column." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }
        if (-Not ($csvData | Get-Member -Name Termination_Date)) {
            Write-Warning "Warning: The CSV file does not contain a Termination_Date column. The termination date will not be included in the output."
        }

        # Extract all employee IDs and Termination Dates into separate arrays
        $employeeIDs = $csvData | Select-Object -ExpandProperty Employee_ID
        $terminationDates = $csvData | Select-Object -ExpandProperty Termination_Date

        $results = @()

        Write-Host "Searching Active Directory for users with matching Employee IDs...`n"

        # Loop through using index so that termination date matches the order of Employee_IDs
        for ($i = 0; $i -lt $employeeIDs.Count; $i++) {
            $employeeID = $employeeIDs[$i]
            $terminationDate = $terminationDates[$i]
        
            # Set termination date to "N/A" if it is null or empty
            if (-not $terminationDate)
            {
                $terminationDate = "N/A"
            }
        
            $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, MemberOf
        
            if ($user) {
                $userGroups = $user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
                $smsUsersMember = $userGroups -contains "SMS Users"
                $sugarbushMember = $userGroups -contains "Sugarbush-SUG-RTP"
        
                # Include users who are members of at least one of the target groups or have an active account
                if ($smsUsersMember -or $sugarbushMember -or $user.Enabled) {
                    # Build the custom object including termination date
                    $result = [PSCustomObject]@{
                        Username           = $user.SamAccountName
                        EmployeeID         = $user.EmployeeID
                        "Termination Date" = $terminationDate
                        "SMS Users"        = if ($smsUsersMember) { "Yes" } else { "No" }
                        "Sugarbush-SUG-RTP"= if ($sugarbushMember) { "Yes" } else { "No" }
                        "Active Account"   = if ($user.Enabled) { "Yes" } else { "No" }
                    }
                    $results += $result
        
                    Write-Host "Username: $($result.Username), EmployeeID: $($result.EmployeeID), Termination Date: $($result.'Termination Date'), SMS Users: $($result.'SMS Users'), Sugarbush-SUG-RTP: $($result.'Sugarbush-SUG-RTP'), Active Account: $($result.'Active Account')"
                }
            }
        }

        # Display results in a table format
        Write-Host "`nResults:"
        $results | Format-Table -AutoSize
    }

    # Prompt to re-run or exit
    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")
