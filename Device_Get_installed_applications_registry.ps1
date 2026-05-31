# Get installed applications from the registry
$installedApps = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    , "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, UninstallString

# Filter out applications without a display name
$installedApps = $installedApps | Where-Object { $_.DisplayName -ne $null }

$FormatEnumerationLimit=-1


# Output the list of installed applications
$installedApps | Format-Table -AutoSize

$installedApps | export-csv installedapps.csv -NoTypeInformation
