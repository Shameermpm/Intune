
# ==========================================
# Check Microsoft.Graph module
# ==========================================
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    Write-Host "Microsoft.Graph module installed successfully." -ForegroundColor Green
}
else {
    Write-Host "Microsoft.Graph module is already installed." -ForegroundColor Green
}

# ==========================================
# Import module & connect
# ==========================================
Import-Module Microsoft.Graph.DeviceManagement
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
setx MSAL_FORCE_WAM 0
Connect-MgGraph -Scopes DeviceManagementManagedDevices.Read.All


# =========================================================================================================#

# ==========================================
# Script start time
# ==========================================
$ScriptStartTime = Get-Date

# ==========================================
# User Input 
# ==========================================
$AppName        = "edge"                # Application Nmae
$Platform       = "windows"                      # Enter the Platform e.g- Windows,AndroidFullyManagedDedicated,Other,AndroidWorkProfile,IOS,AndroidDeviceAdministrator
$FilterOperator = "like"                         # Choose filter operator: 'like' or 'eq'
$OutputPath     = "C:\Temp\DiscoveredApps"       # Application reporting folder Location

# ==========================================
# Ensure output path exists
# ==========================================
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}



# ==========================================
# Validate filter operator
# ==========================================
if ($FilterOperator -notin @("like","eq")) {
    Write-Host "Invalid FilterOperator. Use 'like' or 'eq'." -ForegroundColor Red
    Disconnect-MgGraph
    return
}
# =============================
# GET DETECTED APP COUNT
# =============================
Write-Host "Retrieving detected app count..." -ForegroundColor Yellow
$CountResponse = Invoke-MgGraphRequest `
    -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?`$count=true" `
    -Headers @{ ConsistencyLevel = "eventual" }

$totalDetectedApps = $CountResponse.'@odata.count'
Write-Host "Total detected apps in tenant: $totalDetectedApps" -ForegroundColor Green

# ==========================================
# Get all detected apps
# ==========================================
Write-Host "Retrieving detected apps from Intune..." -ForegroundColor Yellow
$allApps = Get-MgDeviceManagementDetectedApp -All |
    Where-Object { $_.Platform -eq "windows" }

# ==========================================
# Filter apps
# ==========================================
if ($FilterOperator -eq 'like') {

    Write-Host "Applying LIKE filter..." -ForegroundColor Cyan
    $apps = $allApps | Where-Object {
        $_.DisplayName -like "*$AppName*" -and $_.Platform -eq $Platform
    }

}
elseif ($FilterOperator -eq 'eq') {

    Write-Host "Applying EQ filter..." -ForegroundColor Cyan
    $apps = $allApps | Where-Object {
        $_.DisplayName -eq $AppName -and $_.Platform -eq $Platform
    }
}

if (-not $apps) {
    Write-Host "No detected apps found for '$AppName' on platform '$Platform'." -ForegroundColor Red
    #Disconnect-MgGraph
    return
}

# ==========================================
# Collect device mappings
# ==========================================
$result = @()

foreach ($app in $apps) {

    Write-Host "Processing app: $($app.DisplayName) | Version: $($app.Version)" -ForegroundColor Yellow

    $devices = Get-MgDeviceManagementDetectedAppManagedDevice `
        -DetectedAppId $app.Id `
        -All

    if ($devices) {
        foreach ($device in $devices) {
            $result += [PSCustomObject]@{
                DeviceName        = $device.DeviceName
                UserEmail         = $device.EmailAddress
                OsVersion         = $device.Osversion
                AppName           = $app.DisplayName
                AppVersion        = $app.Version
                Platform          = $app.Platform
                Publisher         = $app.Publisher
            }
        }
    }
}

# ==========================================
# Export to CSV
# ==========================================
$FileName = Join-Path $OutputPath (
    "{0}_{1}_{2}.csv" -f (
        ($AppName -replace '\s','_'),
        $Platform,
        (Get-Date -Format 'yyyyMMdd_HHmmss')
    )
)

$result |
    Sort-Object DeviceName, AppName, AppVersion |
    Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8

Write-Host "Report saved to: $FileName" -ForegroundColor Green
Invoke-Item -Path $FileName

# ==========================================
# Script end time & duration
# ==========================================
$ScriptEndTime = Get-Date
$Duration = New-TimeSpan -Start $ScriptStartTime -End $ScriptEndTime

Write-Host (
    "Script execution completed in: {0:hh\:mm\:ss}" -f $Duration
) -ForegroundColor Cyan

# ==========================================
# Disconnect
# ==========================================
#Disconnect-MgGraph
