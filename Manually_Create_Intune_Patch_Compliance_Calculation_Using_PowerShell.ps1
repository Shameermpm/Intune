# Connect to Graph with Intune permissions
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

# Get only Windows devices
$DeviceList = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'Windows'"

# Export to CSV for later processing
$DeviceCsvPath = "C:\TEMP\IntunePatchingReport\IntuneWindowsDevices.csv"
$DeviceList | Select-Object id, deviceName, operatingSystem, complianceState, userPrincipalName,
    serialNumber, osVersion, ownership, joinType, manufacturer, model, skuFamily,
    totalStorageSpaceInBytes, freeStorageSpaceInBytes, lastSyncDateTime |
    Export-Csv -Path $DeviceCsvPath -NoTypeInformation

Write-Host "Device dump exported to $DeviceCsvPath" -ForegroundColor Yellow

# If you want to re-import as CSV for Phase-2 processing:
$DevicesInfos = Import-Csv -Path $DeviceCsvPath
