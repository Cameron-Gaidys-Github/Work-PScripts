# === USER CONFIGURATION ===
$SiteCode = "ENT"  # <- Your SCCM Site Code
$TargetBIOSVersion = [version]"1.43.0"
$CollectionName = "Devices with BIOS >= 1.43.0"
$LimitingCollectionName = "All Systems"

# === IMPORT CONFIGMGR MODULE & CONNECT TO SITE ===
Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
Set-Location "${SiteCode}:"

Write-Host "`n⏳ Querying SCCM for BIOS versions..." -ForegroundColor Cyan

# === GET ALL DEVICES ===
$devices = Get-CMDevice

# === COLLECT DEVICES WITH BIOS >= TARGET ===
$upToDateDevices = @()

foreach ($device in $devices) {
    try {
        $biosInfo = Get-CMBIOSInformation -Device $device -ErrorAction Stop
        if ($biosInfo -and [version]$biosInfo.SMBIOSBIOSVersion -ge $TargetBIOSVersion) {
            $device | Add-Member -MemberType NoteProperty -Name BIOSVersion -Value $biosInfo.SMBIOSBIOSVersion
            $upToDateDevices += $device
        }
    } catch {
        # Skip if BIOS info not available
    }
}

# === SHOW RESULTS ===
if ($upToDateDevices.Count -eq 0) {
    Write-Host "`n❌ No devices found with BIOS >= $TargetBIOSVersion" -ForegroundColor Red
    return
}

Write-Host "`n✅ Devices with BIOS >= $($TargetBIOSVersion):" -ForegroundColor Green
$upToDateDevices | Select-Object Name, BIOSVersion | Format-Table -AutoSize

# === PROMPT TO CREATE COLLECTION ===
$confirmation = Read-Host "`n🛑 Do you want to create collection '$CollectionName' with these devices? (Y/N)"
if ($confirmation -ne "Y") {
    Write-Host "⚠️ Operation cancelled by user." -ForegroundColor Yellow
    return
}

# === CREATE THE COLLECTION ===
New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollectionName -RefreshSchedule (New-CMSchedule -RecurInterval Days -RecurCount 1) | Out-Null

# === ADD DEVICES TO COLLECTION ===
foreach ($device in $upToDateDevices) {
    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceID $device.ResourceID
}

Write-Host "`n🎯 Collection '$CollectionName' created with $($upToDateDevices.Count) devices." -ForegroundColor Green
