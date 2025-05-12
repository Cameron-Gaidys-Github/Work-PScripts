# Script Title: Active Directory Manager Update Script for Job Changes
# 
# Input: Workday Job Change Report CSV file with columns: Employee_ID, New_Manager, Effective_Date
#
# Description:
# This PowerShell script is designed to verify and update the manager-employee relationships in Active Directory (AD) 
# based on job change data provided in a CSV file. It ensures that the `Manager` property in AD is updated to reflect 
# the new manager specified in the CSV file. The script generates a detailed report of current and expected manager 
# relationships and provides options to:
# 
# - Print all users with mismatched managers.
# - Update the `Manager` property in AD for users with mismatched managers.
# 
# Key Features:
# - Prompts the user to relaunch the script with administrator privileges if not already running as admin.
# - Allows the user to select a CSV file containing job change data via a file dialog or manual input.
# - Converts Excel files (.xlsx) to CSV format if necessary.
# - Validates the input file for required columns: "Employee_ID", "New_Manager", and "Effective_Date".
# - Queries AD for users and compares their current manager with the expected manager.
# - Outputs a detailed report of manager-employee relationships, including mismatches.
# - Provides an option to update the `Manager` property in AD for users with mismatched managers.
param (
    [string]$csvFilePath
)

# Prompt user to optionally run as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    $response = Read-Host "This script is not running as Administrator. Would you like to relaunch it with admin rights? Needed to Updated Manager. (Y/N)"
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

# Flashy Title
$asciiArt = @'
         _____                    _____                    _____                    _____                    _____          
         /\    \                  /\    \                  /\    \                  /\    \                  /\    \         
        /::\    \                /::\    \                /::\____\                /::\____\                /::\    \        
       /::::\    \              /::::\    \              /::::|   |               /:::/    /               /::::\    \       
      /::::::\    \            /::::::\    \            /:::::|   |              /:::/    /               /::::::\    \      
     /:::/\:::\    \          /:::/\:::\    \          /::::::|   |             /:::/    /               /:::/\:::\    \     
    /:::/__\:::\    \        /:::/  \:::\    \        /:::/|::|   |            /:::/    /               /:::/__\:::\    \    
   /::::\   \:::\    \      /:::/    \:::\    \      /:::/ |::|   |           /:::/    /                \:::\   \:::\    \   
  /::::::\   \:::\    \    /:::/    / \:::\    \    /:::/  |::|___|______    /:::/    /      _____    ___\:::\   \:::\    \  
 /:::/\:::\   \:::\    \  /:::/    /   \:::\ ___\  /:::/   |::::::::\    \  /:::/____/      /\    \  /\   \:::\   \:::\    \ 
/:::/  \:::\   \:::\____\/:::/____/     \:::|    |/:::/    |:::::::::\____\|:::|    /      /::\____\/::\   \:::\   \:::\____\
\::/    \:::\  /:::/    /\:::\    \     /:::|____|\::/    / ~~~~~/:::/    /|:::|____\     /:::/    /\:::\   \:::\   \::/    /
 \/____/ \:::\/:::/    /  \:::\    \   /:::/    /  \/____/      /:::/    /  \:::\    \   /:::/    /  \:::\   \:::\   \/____/ 
          \::::::/    /    \:::\    \ /:::/    /               /:::/    /    \:::\    \ /:::/    /    \:::\   \:::\    \     
           \::::/    /      \:::\    /:::/    /               /:::/    /      \:::\    /:::/    /      \:::\   \:::\____\    
           /:::/    /        \:::\  /:::/    /               /:::/    /        \:::\__/:::/    /        \:::\  /:::/    /    
          /:::/    /          \:::\/:::/    /               /:::/    /          \::::::::/    /          \:::\/:::/    /     
         /:::/    /            \::::::/    /               /:::/    /            \::::::/    /            \::::::/    /      
        /:::/    /              \::::/    /               /:::/    /              \::::/    /              \::::/    /       
        \::/    /                \::/____/                \::/    /                \::/____/                \::/    /        
         \/____/                  ~~                       \/____/                  ~~                       \/____/        
'@ -split "`n"

# Animate line-by-line with a short delay
foreach ($line in $asciiArt) {
    Write-Host $line -ForegroundColor Cyan
    Start-Sleep -Milliseconds 50
}

do {
    # Clear all variables to reset the script
    Clear-Variable -Name csvFilePath -ErrorAction SilentlyContinue
    Clear-Variable -Name csvData -ErrorAction SilentlyContinue
    Clear-Variable -Name employeeIDs -ErrorAction SilentlyContinue
    Clear-Variable -Name newManagers -ErrorAction SilentlyContinue
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
        $openwindowbool = Read-Host "Would you like to open the file through file explorer? (Y/N)"
        if ($openwindowbool -eq "Y") {
            $csvFilePath = getCSVFile
        } else {
            # Prompt for CSV if not passed as argument (drag-and-drop or manual input)
            $csvFilePath = Read-Host "Enter the full path to the CSV file (drag-and-drop or type manually)"
        }

        # Remove all double-quote and single-quote characters from the file path
        $csvFilePath = $csvFilePath -replace '[\"'']', ''

        # Read the CSV file into $csvData
        $csvData = Import-Csv -Path $csvFilePath

        # Validate path
        # Check if CSV has data and required columns
        if (-not $csvData -or $csvData.Count -eq 0) {
            Write-Host "No data found in the CSV file." -ForegroundColor Yellow
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        if (-not ($csvData | Get-Member -Name Employee_ID -ErrorAction SilentlyContinue)) {
            Write-Host "Error: The CSV file does not contain the required 'Employee_ID' column." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        # Also optional: Check if Employee_ID column has ANY non-blank values
        if (-not ($csvData.Employee_ID | Where-Object { $_ -ne $null -and $_.Trim() -ne "" })) {
            Write-Host "No Employee_ID data found in the CSV file." -ForegroundColor Yellow
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        $employeeIDs = @($csvData | Select-Object -ExpandProperty Employee_ID)
        $newManagers = @($csvData | Select-Object -ExpandProperty New_Manager)
        $effectiveDates = @($csvData | Select-Object -ExpandProperty Effective_Date) # Extract Effective_Date column

        $results = @()

        # Querying Active Directory for users
        Write-Host "Step 5: Querying Active Directory for users..."
        for ($i = 0; $i -lt $employeeIDs.Count; $i++) {
            $employeeID = $employeeIDs[$i]
            $newManagerRaw = $newManagers[$i]
            $effectiveDate = $effectiveDates[$i] # Get the Effective_Date for the current user

            # Parse New Manager Name and Employee ID
            if ($newManagerRaw -match "^(.*?)\s*(?:\(.*?\))?\s*\((\d+)\)$") {
                $newManagerName = $matches[1].Trim() # Extract name and trim any extra spaces
                $newManagerID = $matches[2].Trim()   # Extract Employee ID
            } else {
                $newManagerName = "Invalid Format"
                $newManagerID = "N/A"
            }

            $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, Manager, MemberOf

            if ($user) {
                # Get the current manager's name and Employee ID if the Manager property is populated
                if ($user.Manager) {
                    $currentManager = Get-ADUser -Identity $user.Manager -Properties Name, EmployeeID
                    $currentManagerName = $currentManager.Name -replace "\s+\(.*\)$", "" # Remove suffix like (SUG) or (On Leave)
                    $currentManagerID = $currentManager.EmployeeID
                } else {
                    $currentManagerName = "No Manager Assigned"
                    $currentManagerID = "N/A"
                }

                # Determine account status
                $status = if ($user.Enabled) { "Active" } else { "Inactive" }

                # Compare Current Manager ID and New Manager ID
                $managerMatch = if ($currentManagerID -eq $newManagerID) { "Yes" } else { "No" }

                # Check group membership
                $userGroups = $user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
                $isSMSUser = $userGroups -contains "SMS Users"
                $isSugarbushRTP = $userGroups -contains "Sugarbush-SUG-RTP"

                # Add user details to results
                $results += [PSCustomObject]@{
                    Name                 = $user.Name -replace "\s+\(.*\)$", "" # Remove suffix like (SUG) or (On Leave)
                    Username             = $user.SamAccountName
                    EmployeeID           = $user.EmployeeID
                    Status               = $status
                    CurrentManager       = $currentManagerName
                    CurrentManagerID     = $currentManagerID
                    NewManager           = $newManagerName
                    NewManagerID         = $newManagerID
                    ManagerMatch         = $managerMatch
                    EffectiveDate        = $effectiveDate
                    SMSUsers             = if ($isSMSUser) { "Yes" } else { "No" }
                    SugarbushSUGRTP      = if ($isSugarbushRTP) { "Yes" } else { "No" }
                }
            } else {
                Write-Host "No user found for Employee ID: $employeeID" -ForegroundColor Yellow
            }
        }

        # Display results in a table format
        Write-Host "`nResults:"
        $results | Format-Table @{Label="Name"; Expression={"{0}" -f $_.Name}}, 
                                    @{Label="Username"; Expression={"{0}" -f $_.Username}},
                                    @{Label="Employee ID"; Expression={"{0}" -f $_.EmployeeID}},
                                    @{Label="Status"; Expression={"{0}" -f $_.Status}},
                                    @{Label="Current Manager"; Expression={"{0}" -f $_.CurrentManager}},
                                    @{Label="Current Manager ID"; Expression={"{0}" -f $_.CurrentManagerID}},
                                    @{Label="New Manager"; Expression={"{0}" -f $_.NewManager}},
                                    @{Label="New Manager ID"; Expression={"{0}" -f $_.NewManagerID}},
                                    @{Label="Manager Match"; Expression={"{0}" -f $_.ManagerMatch}},
                                    @{Label="Effective Date"; Expression={"{0}" -f $_.EffectiveDate}},
                                    @{Label="SMS Users"; Expression={"{0}" -f $_.SMSUsers}},
                                    @{Label="Sugarbush-SUG-RTP"; Expression={"{0}" -f $_.SugarbushSUGRTP}} -AutoSize | Out-String | Write-Host

        # Prompt user for confirmation before updating managers
        $updateConfirmation = Read-Host "`nDo you want to update managers in Active Directory? Type 'Yes' to proceed or anything else to cancel"
        if ($updateConfirmation -eq "Yes") {
            Write-Host "Updating managers in Active Directory..." -ForegroundColor Cyan
            foreach ($result in $results) {
                if ($result.ManagerMatch -eq "No" -and $result.NewManagerID -ne "N/A" -and $result.NewManagerID.Trim() -ne "") {
                    try {
                        Write-Host "Processing user: $($result.Name) with EmployeeID: $($result.EmployeeID)" -ForegroundColor Yellow
                        Write-Host "New Manager ID: $($result.NewManagerID)" -ForegroundColor Yellow
        
                        $userToUpdate = Get-ADUser -LDAPFilter "(employeeID=$($result.EmployeeID))" -Properties DistinguishedName
                        $newManager   = Get-ADUser -LDAPFilter "(employeeID=$($result.NewManagerID))" -Properties DistinguishedName
        
                        if ($userToUpdate -and $newManager) {
                            Write-Host "Updating manager for user '$($userToUpdate.SamAccountName)' to '$($newManager.SamAccountName)'" -ForegroundColor Cyan
                            Write-Host "New Manager DN: $($newManager.DistinguishedName)" -ForegroundColor Cyan
        
                            Set-ADUser -Identity $userToUpdate.DistinguishedName -Manager $newManager.DistinguishedName
                            Write-Host "✅ Updated manager for user '$($userToUpdate.SamAccountName)' to '$($newManager.SamAccountName)'" -ForegroundColor Green
                        } elseif (-not $userToUpdate) {
                            Write-Host "❌ User with Employee ID $($result.EmployeeID) not found." -ForegroundColor Yellow
                        } elseif (-not $newManager) {
                            Write-Host "❌ New manager with Employee ID $($result.NewManagerID) not found." -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "❌ Failed to update manager for user $($result.Name): $_" -ForegroundColor Red
                    }
                }
            }      
        
        } else {
            Write-Host "Manager updates canceled by user." -ForegroundColor Yellow
        }
    }

    # Prompt to re-run or exit
    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")