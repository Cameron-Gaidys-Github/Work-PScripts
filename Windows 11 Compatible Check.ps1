# TPM Check
try {
    $tpm = Get-WmiObject -Namespace "Root\CIMV2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction Stop
    $tpmReady = if ($tpm.SpecVersion -like "*2.0*") { "✅ Pass" } else { "❌ TPM not 2.0" }
} catch {
    $tpmReady = "❌ TPM not found"
}

# Secure Boot Check
try {
    $secureBoot = if (Confirm-SecureBootUEFI) { "✅ Pass" } else { "❌ Disabled or unsupported" }
} catch {
    $secureBoot = "❌ Unsupported (Legacy BIOS or non-UEFI)"
}

# CPU Check
$cpu = Get-CimInstance Win32_Processor
$cpuCores = $cpu.NumberOfCores
$cpuSpeedGHz = [math]::Round($cpu.MaxClockSpeed / 1000, 2)
$cpuStatus = if ($cpuCores -ge 2 -and $cpuSpeedGHz -ge 1.0) { "✅ Pass" } else { "❌ $cpuCores cores @ $cpuSpeedGHz GHz" }

# RAM Check
$ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$ramStatus = if ($ramGB -ge 4) { "✅ Pass" } else { "❌ $ramGB GB" }

# Storage Check
$storageGB = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").Size / 1GB, 2)
$storageStatus = if ($storageGB -ge 64) { "✅ Pass" } else { "❌ $storageGB GB" }

# Output results
[PSCustomObject]@{
    TPM          = $tpmReady
    SecureBoot   = $secureBoot
    CPU          = $cpuStatus
    RAM_GB       = "$ramGB GB ($ramStatus)"
    Storage_GB   = "$storageGB GB ($storageStatus)"
}
