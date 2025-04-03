param (
    [string]$csvFilePath
)

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
try
{
    Import-Module ActiveDirectory -ErrorAction Stop
} 
catch 
{
    Write-Host "Error: Active Directory module is not available. Ensure RSAT is installed." -ForegroundColor Red
    Read-Host -Prompt "Press Enter to close this window"
    exit
}

# Verify the module is imported
Write-Host "Step 3: Verifying module import..."
if (Get-Module -Name ActiveDirectory) {
    Write-Host "Step 4: Specify the CSV file to import..."

    # Prompt for CSV if not passed as argument (i.e. not drag-and-drop)
    if (-not $csvFilePath) {
        $csvFilePath = Read-Host "Enter the full path to the CSV file"
    }

    # Remove all double-quote characters from the file path
    $csvFilePath = $csvFilePath -replace '"', ''

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

    $employeeIDs = $csvData | Select-Object -ExpandProperty Employee_ID
    $array = @($employeeIDs)

    $results = @()

    Write-Host "Searching Active Directory for users with matching Employee IDs...`n"
    foreach ($employeeID in $array) {
    $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, SamAccountName, Name, Enabled, MemberOf

    if ($user) {
        $userGroups = $user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
        $smsUsersMember = $userGroups -contains "SMS Users"
        $sugarbushMember = $userGroups -contains "Sugarbush-SUG-RTP"

        # Include users who are members of at least one of the target groups or have an active account
        if ($smsUsersMember -or $sugarbushMember -or $user.Enabled) {
            # Add the result to the array
            $result = [PSCustomObject]@{
                Username          = $user.SamAccountName
                EmployeeID        = $user.EmployeeID
                "SMS Users"       = if ($smsUsersMember) { "Yes" } else { "No" }
                "Sugarbush-SUG-RTP" = if ($sugarbushMember) { "Yes" } else { "No" }
                "Active Account"  = if ($user.Enabled) { "Yes" } else { "No" }
            }
            $results += $result

            # Output the result to the terminal
            Write-Host "Username: $($result.Username), EmployeeID: $($result.EmployeeID), SMS Users: $($result.'SMS Users'), Sugarbush-SUG-RTP: $($result.'Sugarbush-SUG-RTP'), Active Account: $($result.'Active Account')"
        }
    }
}
    Read-Host -Prompt "`nPress Enter to close this window"
} 
        
    # End of script
