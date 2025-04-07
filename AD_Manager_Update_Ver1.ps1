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

        if (-Not ($csvData | Get-Member -Name Employee_ID)) {
            Write-Host "Error: The CSV file does not contain the required 'Employee_ID' column." -ForegroundColor Red
            Read-Host -Prompt "Press Enter to close this window"
            exit
        }

        $employeeIDs = @($csvData | Select-Object -ExpandProperty Employee_ID)
        $newManagers = @($csvData | Select-Object -ExpandProperty New_Manager)

        $results = @()

        # Querying Active Directory for users
        Write-Host "Step 5: Querying Active Directory for users..."
        for ($i = 0; $i -lt $employeeIDs.Count; $i++) {
            $employeeID = $employeeIDs[$i]
            $newManagerRaw = $newManagers[$i]

            # Parse New Manager Name and Employee ID
            if ($newManagerRaw -match "^(.*?)\s*(?:\(.*?\))?\s*\((\d+)\)$") {
                $newManagerName = $matches[1].Trim() # Extract name and trim any extra spaces
                $newManagerID = $matches[2].Trim()   # Extract Employee ID
            } else {
                $newManagerName = "Invalid Format"
                $newManagerID = "N/A"
            }

            $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, Manager

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
                                @{Label="Manager Match"; Expression={"{0}" -f $_.ManagerMatch}} -AutoSize
    }

    # Prompt to re-run or exit
    $choice = Read-Host "`nPress R to re-run the script or Enter to close"
} while ($choice -eq "R")