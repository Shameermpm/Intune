# Connect to Graph with Intune permissions
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

# Get only Windows devices
$DeviceList = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'Windows'"

# Define output folder
$WorkingFolder = "C:\TEMP\IntunePatchingReport"
New-Item -ItemType Directory -Path $WorkingFolder -Force | Out-Null

# Export to CSV with correct columns
$DeviceCsvPath = "$WorkingFolder\IntuneWindowsDevices.csv"
$DeviceList | Select-Object `
    Id,
    DeviceName,
    ManagedDeviceName,
    OperatingSystem,
    OSVersion,
    ComplianceState,
    DeviceEnrollmentType,
    DeviceRegistrationState,
    ManagementAgent,
    ManagementState,
    Manufacturer,
    Model,
    SerialNumber,
    UserDisplayName,
    UserPrincipalName,
    EmailAddress,
    EnrolledDateTime,
    LastSyncDateTime,
    TotalStorageSpaceInBytes,
    FreeStorageSpaceInBytes,
    ManagedDeviceOwnerType,
    DeviceCategoryDisplayName |
    Export-Csv -Path $DeviceCsvPath -NoTypeInformation

Write-Host "Device dump exported to $DeviceCsvPath" -ForegroundColor Yellow

# Re-import if you want to process as CSV later
$DevicesInfos = Import-Csv -Path $DeviceCsvPath
