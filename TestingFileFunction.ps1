# Function to open file explorer and select a file
function getCSVFile {
    # Load the required assembly for Windows Forms
    Add-Type -AssemblyName System.Windows.Forms

    # File Selection window
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('UserProfile') + "\Downloads"
        Filter = 'Spreadsheet (*.csv, *.xlsx)|*.csv;*.xlsx'
        Title = "Select Term Report"
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

# Main script logic
do {
    # Prompt user to select a file using the getCSVFile function
    Write-Host "Step 4: Specify the CSV file to import..."
    $csvFilePath = getCSVFile

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

    # Rest of your script logic...
    $choice = Read-Host "`nPress R to re-run the script, D to display a message, or Enter to close"
    if ($choice -eq "D") {
        Write-Host "`nYou selected D! Here's your custom message:" -ForegroundColor Green
        Write-Host "Active Directory Manager Update Script is running smoothly!" -ForegroundColor Cyan
    }
} while ($choice -eq "R")