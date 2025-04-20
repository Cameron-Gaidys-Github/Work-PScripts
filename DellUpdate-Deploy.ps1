function Prompt-For-Hostnames {
    Write-Host "Would you like to open file through file explorer? (Y/N): " -NoNewline
    $response = Read-Host

    if ($response -match '^(Y|y)$') {
        Add-Type -AssemblyName System.Windows.Forms
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = "CSV files (*.csv)|*.csv"
        if ($openFile.ShowDialog() -eq 'OK') {
            $csvPath = $openFile.FileName
            try {
                $csvData = Import-Csv $csvPath
                if ($csvData -and $csvData[0].Hostname) {
                    return $csvData.Hostname
                } else {
                    Write-Warning "CSV is missing a 'Hostname' column."
                    return @()
                }
            } catch {
                Write-Error "Failed to load CSV: $_"
                return @()
            }
        }
    } else {
        Write-Host "Enter hostnames separated by commas (e.g. PC-01,PC-02):"
        $input = Read-Host
        return $input -split ',' | ForEach-Object { $_.Trim() }
    }
}

function Install-DCU {
    param ($session)
    Invoke-Command -Session $session -ScriptBlock {
        $dcuPath = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
        $exePath = "$env:TEMP\DellCommandUpdate.exe"
        $url = "https://downloads.dell.com/FOLDER12345678M/1/Dell-Command-Update-Application_XXXX.exe"

        if (-not (Test-Path $dcuPath)) {
            Write-Output "Installing Dell Command | Update..."
            Invoke-WebRequest -Uri $url -OutFile $exePath -UseBasicParsing
            Start-Process -FilePath $exePath -ArgumentList "/quiet" -Wait
        }
    }
}

function Run-DellUpdates {
    param ($session, $hostname)
    Invoke-Command -Session $session -ScriptBlock {
        $dcuPath = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
        if (Test-Path $dcuPath) {
            Write-Output "[$env:COMPUTERNAME] Running Dell updates..."
            $result = Start-Process -FilePath $dcuPath -ArgumentList "/applyUpdates", "/silent", "/reboot=disable" -Wait -PassThru
            if ($result.ExitCode -eq 0) {
                Write-Output "[$env:COMPUTERNAME] Update complete."
            } else {
                Write-Output "[$env:COMPUTERNAME] Update failed. Exit Code: $($result.ExitCode)"
            }
        } else {
            Write-Output "[$env:COMPUTERNAME] DCU not found."
        }
    }
}

function Run-UpdateCycle {
    $hostnames = Prompt-For-Hostnames

    if (-not $hostnames -or $hostnames.Count -eq 0) {
        Write-Warning "No hostnames provided."
        return
    }

    foreach ($computerName in $hostnames) {
        if (-not (Get-ADComputer -Filter { Name -eq $computerName } -ErrorAction SilentlyContinue)) {
            Write-Warning "❌ $computerName not found in AD."
            continue
        }

        Write-Host "`n➡️ Connecting to $computerName..."
        try {
            $session = New-PSSession -ComputerName $computerName -ErrorAction Stop

            Install-DCU -session $session
            Run-DellUpdates -session $session -hostname $computerName

            Remove-PSSession $session
        } catch {
            Write-Warning "⚠️ Failed to connect to ${computerName}: $_"
        }
    }
}

# Ensure AD module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "❌ Active Directory module missing. Run as Admin with RSAT tools."
    exit
}

# === MAIN LOOP ===
do {
    Run-UpdateCycle
    Write-Host "`nPress [Enter] to exit or [R] to rerun." -ForegroundColor Yellow
    $choice = Read-Host
} while ($choice -match '^(R|r)$')
