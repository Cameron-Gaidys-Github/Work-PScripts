# Leave Script: Active Directory Group Membership Verification for Leave Reports
# 
# Input: Workday Leave Report CSV file (Employee_ID, Leave_Started, Estimated_Return, Actual_End_Date, Initiated)
#
# This PowerShell script is used to verify Active Directory (AD) group memberships and account statuses 
# for employees listed in a leave report CSV file. It is typically used to track employees on leave 
# and their associated AD account details.

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

    return $FileName
}

do {
    Clear-Variable -Name csvFilePath -ErrorAction SilentlyContinue
    Clear-Variable -Name csvData -ErrorAction SilentlyContinue
    Clear-Variable -Name leaveDetails -ErrorAction SilentlyContinue
    Clear-Variable -Name results -ErrorAction SilentlyContinue

    Write-Host "Step 1: Specify the CSV file to import..."

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

    # Handle empty CSV file
    if (-not $csvData -or $csvData.Count -eq 0) {
        Write-Host "The CSV file is empty. Please provide a file with data." -ForegroundColor Yellow
        $choice = Read-Host "`nPress R to re-run the script or Enter to close"
        if ($choice -eq "R") { continue } else { break }
    }

    if (-Not ($csvData | Get-Member -Name Employee_ID)) {
        Write-Host "Error: The CSV file does not contain the required Employee_ID column." -ForegroundColor Red
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

        $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, MemberOf, Title

        $smsUsersMember = $false
        $sugarbushMember = $false
        $action = "N/A"
        $jobTitle = "N/A"

        if ($user) {
            $userGroups = $user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
            $smsUsersMember = $userGroups -contains "SMS Users"
            $sugarbushMember = $userGroups -contains "Sugarbush-SUG-RTP"

            if ($user.SamAccountName -match '^\d+$') {
                $username = "N/A"
            } else {
                $username = $user.SamAccountName
            }

            # Extract job title and determine action
            $jobTitle = $user.Title
            $jobTitlePrefix = if ($jobTitle -match "^(.*?)[,\s]") { $matches[1].Trim().ToLower() } else { $jobTitle.ToLower() }
            $action = if ($jobTitlePrefix -in @("supervisor", "manager", "director")) {
                "Suspend with Delegation"
            } else {
                "Suspend with no Delegate"
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
            "SMS Users"     = if ($smsUsersMember) { "Yes" } else { "No" }
            "Sugarbush-SUG-RTP" = if ($sugarbushMember) { "Yes" } else { "No" }
            JobTitle        = $jobTitle
            Action          = $action
        }
        $results += $result
    }

    Write-Host "`nResults:"
    Write-Host "Preferred Name                  Username                Employee ID | Active | Leave Started                   Estimated Return                 Actual End Date                  Initiated                         SMS Users | Sugarbush-SUG-RTP | Job Title                 | Action"
    Write-Host "-------------------------------|----------------------|-------------|--------|--------------------------------|--------------------------------|--------------------------------|--------------------------------|-----------|-------------------|---------------------------|-------------------------"

    $results | ForEach-Object {
        $line = "{0,-30} | {1,-20} | {2,-11} | {3,-6} | {4,-30} | {5,-30} | {6,-30} | {7,-30} | {8,-9} | {9,-17} | {10,-25} | {11,-25}" -f `
            ($_.PreferredName.Substring(0, [Math]::Min($_.PreferredName.Length, 30))), `
            ($_.Username.Substring(0, [Math]::Min($_.Username.Length, 20))), `
            $_.EmployeeID, `
            $_.Active, `
            ($_.LeaveStarted.Substring(0, [Math]::Min($_.LeaveStarted.Length, 30))), `
            ($_.EstimatedReturn.Substring(0, [Math]::Min($_.EstimatedReturn.Length, 30))), `
            ($_.ActualEndDate.Substring(0, [Math]::Min($_.ActualEndDate.Length, 30))), `
            ($_.Initiated.Substring(0, [Math]::Min($_.Initiated.Length, 30))), `
            $_.'SMS Users', `
            $_.'Sugarbush-SUG-RTP', `
            ($_.JobTitle.Substring(0, [Math]::Min($_.JobTitle.Length, 25))), `
            $_.Action
        Write-Host $line
    }

    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")