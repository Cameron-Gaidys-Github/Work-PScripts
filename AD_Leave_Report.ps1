param (
    [string]$csvFilePath
)

# Set the PowerShell console window size
[console]::WindowWidth = 220
[console]::WindowHeight = 40

function getCSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select Leave Report"
    }
    $null = $FileBrowser.ShowDialog()
    $FileName = $FileBrowser.FileName

    if ($FileBrowser.FileName -match '.xlsx') {
        $Excel = New-Object -ComObject Excel.Application
        $wb = $Excel.Workbooks.Open($FileBrowser.FileName)
        $FileName = $FileBrowser.FileName.Replace(".xlsx", ".csv")
        foreach ($ws in $wb.Worksheets) {
            $ws.SaveAs($FileName, 6)
        }
        $Excel.Quit()
    }

    $firstline = Get-Content $FileName -First 6
    if ($firstline -match "Terminations") {
        (Get-Content $FileName | Select-Object -Skip 6) | Set-Content $FileName
    }

    return $FileName
}

do {
    Clear-Variable -Name csvFilePath -ErrorAction SilentlyContinue
    Clear-Variable -Name csvData -ErrorAction SilentlyContinue
    Clear-Variable -Name leaveDetails -ErrorAction SilentlyContinue
    Clear-Variable -Name results -ErrorAction SilentlyContinue

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

    Write-Host "Step 2: Importing Active Directory module..."
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Host "Error: Active Directory module is not available. Ensure RSAT is installed." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to close this window"
        exit
    }

    Write-Host "Step 3: Verifying module import..."
    if (Get-Module -Name ActiveDirectory) {
        Write-Host "Step 4: Specify the CSV file to import..."

        $openwindowbool = Read-Host "Would you like to open file through file explorer? (Y/N)"
        if ($openwindowbool -eq "Y") {
            $csvFilePath = getCSVFile
        } else {
            $csvFilePath = Read-Host "Enter the full path to the CSV file (drag-and-drop or type manually)"
        }

        $csvFilePath = $csvFilePath -replace '[\"'']', ''

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

        if (-Not ($csvData | Get-Member -Name Employee_ID)) {
            Write-Host "Error: The CSV file does not contain the required Employee_ID column." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        $leaveDetails = $csvData | Select-Object -Property Employee_ID, Preferred_Name, Leave_Started, Estimated_Return, Actual_End_Date, Initiated
        $results = @()

        Write-Host "Searching Active Directory for users with matching Employee IDs...`n"

        foreach ($leave in $leaveDetails) {
            $employeeID = $leave.Employee_ID
            $preferredName = $leave.Preferred_Name
            $leaveStarted = $leave.Leave_Started
            $estimatedReturn = $leave.Estimated_Return
            $actualEndDate = $leave.Actual_End_Date
            $initiated = $leave.Initiated

            $leaveStarted = if ($leaveStarted) { $leaveStarted } else { "N/A" }
            $estimatedReturn = if ($estimatedReturn) { $estimatedReturn } else { "N/A" }
            $actualEndDate = if ($actualEndDate) { $actualEndDate } else { "N/A" }
            $initiated = if ($initiated) { $initiated } else { "N/A" }

            $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled

            if ($user -and $user.SamAccountName) {
                if ($user.SamAccountName -match '^\d+$') {
                    $username = "N/A"
                } else {
                    $username = $user.SamAccountName
                }
            } else {
                $username = "N/A"
            }

            $result = [PSCustomObject]@{
                PreferredName   = $preferredName
                Username        = $username
                EmployeeID      = $employeeID
                Active          = if ($user -and $user.Enabled) { "Yes" } else { "No" }
                LeaveStarted    = $leaveStarted
                EstimatedReturn = $estimatedReturn
                ActualEndDate   = $actualEndDate
                Initiated       = $initiated
            }
            $results += $result
        }

        Write-Host "`nResults:"
        Write-Host "Preferred Name                  Username                Employee ID | Active | Leave Started                   Estimated Return                 Actual End Date                  Initiated"
        Write-Host "-------------------------------|----------------------|-------------|--------|--------------------------------|--------------------------------|--------------------------------|----------------"

        $results | ForEach-Object {
            $line = "{0,-30} | {1,-20} | {2,-11} | {3,-6} | {4,-30} | {5,-30} | {6,-30} | {7,-30}" -f `
                ($_.PreferredName.Substring(0, [Math]::Min($_.PreferredName.Length, 30))), `
                ($_.Username.Substring(0, [Math]::Min($_.Username.Length, 20))), `
                $_.EmployeeID, `
                $_.Active, `
                ($_.LeaveStarted.Substring(0, [Math]::Min($_.LeaveStarted.Length, 30))), `
                ($_.EstimatedReturn.Substring(0, [Math]::Min($_.EstimatedReturn.Length, 30))), `
                ($_.ActualEndDate.Substring(0, [Math]::Min($_.ActualEndDate.Length, 30))), `
                ($_.Initiated.Substring(0, [Math]::Min($_.Initiated.Length, 30)))
            Write-Host $line
        }
    }

    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")