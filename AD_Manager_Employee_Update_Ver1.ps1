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

    # Display results in a neatly aligned bar-separated format
    Write-Host "`nResults:"
    Write-Host "Name                             Username             Employee ID | Current Manager                  Current Manager ID | Expected Manager                Expected Manager ID | Manager Match"
    Write-Host "------------------------------------------------------------------|-----------------------------------------------------|-----------------------------------------------------|--------------"

    $results | ForEach-Object {
        $line = "{0,-30}  {1,-20}  {2,-11} | {3,-30} {4,-20} | {5,-30} {6,-20} | {7,-10}" -f `
            ($_.Name.Substring(0, [Math]::Min($_.Name.Length, 30))), `
            ($_.Username.Substring(0, [Math]::Min($_.Username.Length, 20))), `
            $_.EmployeeID, `
            ($_.CurrentManager.Substring(0, [Math]::Min($_.CurrentManager.Length, 30))), `
            $_.CurrentManagerID, `
            ($_.ExpectedManager.Substring(0, [Math]::Min($_.ExpectedManager.Length, 30))), `
            $_.ExpectedManagerID, `
            $_.ManagerMatch
        Write-Host $line
    }

    # Ask the user if they want to display only users with non-matching managers
    $filterChoice = Read-Host "`nWould you like to display only users with non-matching managers? (Y/N)"
    if ($filterChoice -eq "Y") {
        Write-Host "`nUsers with non-matching managers:"
        Write-Host "Name                             Username             Employee ID | Current Manager                  Current Manager ID | Expected Manager                Expected Manager ID | Manager Match"
        Write-Host "------------------------------------------------------------------|-----------------------------------------------------|-----------------------------------------------------|--------------"

        $results | Where-Object { $_.ManagerMatch -eq "No" } | ForEach-Object {
            $line = "{0,-30}  {1,-20}  {2,-11} | {3,-30} {4,-20} | {5,-30} {6,-20} | {7,-10}" -f `
                ($_.Name.Substring(0, [Math]::Min($_.Name.Length, 30))), `
                ($_.Username.Substring(0, [Math]::Min($_.Username.Length, 20))), `
                $_.EmployeeID, `
                ($_.CurrentManager.Substring(0, [Math]::Min($_.CurrentManager.Length, 30))), `
                $_.CurrentManagerID, `
                ($_.ExpectedManager.Substring(0, [Math]::Min($_.ExpectedManager.Length, 30))), `
                $_.ExpectedManagerID, `
                $_.ManagerMatch
            Write-Host $line
        }
}
 
     # Prompt to re-run or exit
     $choice = Read-Host "`nPress R to re-run the script or Enter to close"
 } while ($choice -eq "R")