﻿
#<# 
#.SYNOPSIS
#<Intune_Patch_Compliance_Calculation_Manually>
#>

<# 
#.DESCRIPTION
#<Intune_Patch_Compliance_Calculation_Manually>
#>


<# 
#.INPUTS
#<Provide all required information in User Input Section
#>

<# 
#.OUTPUTS
#<You will get Intune_Patch_Compliance report in CSV>
#>

<# 
#.NOTES
#Version:                   v4.0
#Update Date :              16 June 2026



#>
$error.clear()
cls
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
$ErrorActionPreference = 'SilentlyContinue'

#--------------------------------  User Input Section Start -----------------------------------------------------------------

$DeviceList             = "C:\Users\CM\Downloads\Dump.csv" # Location of Intune Device data csv file
$WorkingFolder          = "C:\TEMP\Patching_Compliance_Status"
$OfflineExclusionDays   =  10                                         # <-- Change this per prodcution start day (e.g. 7, 14, 21)
$DevcieNamePrefix       = ("VM","DESKTOP","ANCD")                      #  Device Name Prefix

#-------------------------------- User Input Section End---------------------------------------------------------------------
$startTime = Get-Date
Write-Host "===============================Phase-1 (Exporting Intune Device Dump) ======================================================(Started)" -ForegroundColor Green
$error.clear()
$Path = "$WorkingFolder\PD_Dump\"
New-Item -ItemType Directory -Path "$WorkingFolder\PD_Dump\" -Force | Out-Null

#==============================================================================================================================
# PHASE 1 - MgGraph Authentication
#==============================================================================================================================

$DevicesInfos = Import-Csv -Path $DeviceList |
    Select-Object @{Name="DeviceId"; Expression={$_.("Device ID")}},
                  @{Name="SerialNumber"; Expression={$_.("Serial number")}},
                  @{Name="DeviceName"; Expression={$_.("Device name")}},
                  @{Name="Ownership"; Expression={$_.Ownership}},
                  @{Name="OSVersion"; Expression={$_.("OS version")}},
                  @{Name="Primary user UPN"; Expression={$_.("Primary user UPN")}},
                  @{Name="Last check-in"; Expression={$_.("Last check-in")}},
                  @{Name="JoinType"; Expression={$_.JoinType}},
                  @{Name="Manufacturer"; Expression={$_.Manufacturer}},
                  @{Name="Model"; Expression={$_.Model}},
                  @{Name="Managed by"; Expression={$_.("Managed by")}},
                  @{Name="SkuFamily"; Expression={$_.SkuFamily}},
                  @{Name="Total storage"; Expression={$_.("Total storage")}},
                  @{Name="Free storage"; Expression={$_.("Free storage")}}
Write-Host "" 
write-Host "Importing Intune Device Dump completed " -ForegroundColor Yellow
Write-Host "" 

Write-Host "===============================Phase-1 (Importing Intune Device Dump) ====================================================(Completed)" -ForegroundColor Green
Write-Host ""
Write-Host "===============================Phase-2 (Downloading and Creating MS Patch List) ============================================(Started)" -ForegroundColor Green

$Date                  = Get-Date -Format "MMMMMMMM dd, yyyy"
$OutFileMP             = "$WorkingFolder\MicrosoftPatchList.csv"
$OutFileLP             = "$WorkingFolder\MicrosoftLatestPatchList.csv"
$MergeOverallFile      = "$WorkingFolder\MergeOverallFile.csv"
$Final_Patching_Report = "$WorkingFolder\Final_Patching_Report.csv"
$PatchingMonth         = ""
$PatchReleaseDays      = 0

Write-Host "Auto-scraping Windows Build Version Map from Microsoft Learn..." -ForegroundColor Yellow

$buildInfoArray  = @()
$DynamicBuildMap = @{}   

function Merge-BuildMap {
    param(
        [string]$URL,
        [string]$OSPrefix,
        [ref]$BuildArray,
        [ref]$BuildMap
    )
    try {
        $html = (Invoke-WebRequest -Uri $URL -UseBasicParsing -ErrorAction Stop).Content
        
        $tableMatches = [regex]::Matches($html, '\|\s*(\d{2}H[12]\d*)\s*\|[^|]+\|[^|]+\|[^|]+\|[^|]+\|[^|]+\|\s*(\d{4,5})\.\d+\s*\|')
        foreach ($m in $tableMatches) {
            $ver = $m.Groups[1].Value.Trim(); $build = $m.Groups[2].Value.Trim()
            $label = "$OSPrefix-$ver"
            if (($BuildArray.Value | Select-Object -ExpandProperty Build) -notcontains $build) {
                $o = New-Object PSObject; $o | Add-Member -MemberType NoteProperty -Name Build -Value $build
                $o | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value "$OSPrefix $ver"
                $BuildArray.Value += $o
            }
            if (-not $BuildMap.Value.ContainsKey($build)) { $BuildMap.Value[$build] = $label }
        }
       $headMatches = [regex]::Matches($html, 'Version\s+(\d{2}H[12]\d*)\s*\(OS build\s+(\d{4,5})\)')
        foreach ($m in $headMatches) {
            $ver = $m.Groups[1].Value.Trim(); $build = $m.Groups[2].Value.Trim()
            $label = "$OSPrefix-$ver"
            if (($BuildArray.Value | Select-Object -ExpandProperty Build) -notcontains $build) {
                $o = New-Object PSObject; $o | Add-Member -MemberType NoteProperty -Name Build -Value $build
                $o | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value "$OSPrefix $ver"
                $BuildArray.Value += $o
            }
            if (-not $BuildMap.Value.ContainsKey($build)) { $BuildMap.Value[$build] = $label }
        }
        Write-Host "  Scraped $URL  ->  $($BuildMap.Value.Count) total builds so far" -ForegroundColor Cyan
    } catch {
        Write-Host "  Warning: Could not scrape $URL  ($_)" -ForegroundColor Yellow
    }
}

Merge-BuildMap -URL "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information" -OSPrefix "Win11" -BuildArray ([ref]$buildInfoArray) -BuildMap ([ref]$DynamicBuildMap)
Merge-BuildMap -URL "https://learn.microsoft.com/en-us/windows/release-health/release-information"          -OSPrefix "Win10" -BuildArray ([ref]$buildInfoArray) -BuildMap ([ref]$DynamicBuildMap)

$staticFallback = @(
    @{B="28000";OS="Win11 26H1";L="Win11-26H1"},
    @{B="26200";OS="Win11 25H2";L="Win11-25H2"},
    @{B="26100";OS="Win11 24H2";L="Win11-24H2"},
    @{B="22631";OS="Win11 23H2";L="Win11-23H2"},
    @{B="22621";OS="Win11 22H2";L="Win11-22H2"},
    @{B="22000";OS="Win11 21H2";L="Win11-21H2"},
    @{B="19045";OS="Win10 22H2";L="Win10-22H2"},
    @{B="19044";OS="Win10 21H2";L="Win10-21H2"},
    @{B="19043";OS="Win10 21H1";L="Win10-21H1"},
    @{B="19042";OS="Win10 20H2";L="Win10-20H2"},
    @{B="19041";OS="Win10 2004";L="Win10-2004"},
    @{B="18363";OS="Win10 1909";L="Win10-1909"},
    @{B="18362";OS="Win10 1903";L="Win10-1903"},
    @{B="17763";OS="Win10 1809";L="Win10-1809"},
    @{B="17134";OS="Win10 1803";L="Win10-1803"},
    @{B="16299";OS="Win10 1709";L="Win10-1709"},
    @{B="15063";OS="Win10 1703";L="Win10-1703"},
    @{B="14393";OS="Win10 1607";L="Win10-1607"},
    @{B="10586";OS="Win10 1511";L="Win10-1511"},
    @{B="10240";OS="Win10 1507";L="Win10-1507"},
    @{B="9600" ;OS="Windows 8.1";L="Win8.1"},
    @{B="7601" ;OS="Windows 7"  ;L="Win7"}
)
foreach ($fb in $staticFallback) {
    if (($buildInfoArray | Select-Object -ExpandProperty Build) -notcontains $fb.B) {
        $o = New-Object PSObject
        $o | Add-Member -MemberType NoteProperty -Name Build           -Value $fb.B
        $o | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $fb.OS
        $buildInfoArray += $o
    }
    if (-not $DynamicBuildMap.ContainsKey($fb.B)) { $DynamicBuildMap[$fb.B] = $fb.L }
}
Write-Host "Build map ready. Total known major builds: $($buildInfoArray.Count)" -ForegroundColor Green

#==============================================================================================================================
# HOTPATCH SCRAPING
# URL: https://support.microsoft.com/en-us/help/5068391
#==============================================================================================================================

Write-Host "Scraping Hotpatch builds from Microsoft..." -ForegroundColor Yellow

$HotpatchOSBuilds  = @{}
$HotpatchKBMap     = @{}
$HotpatchDateMap   = @{}
$HotpatchAllBuilds = @()

try {
    $hpHTML = (Invoke-WebRequest -Uri "https://support.microsoft.com/en-us/help/5068391" -UseBasicParsing -ErrorAction Stop).Content
    $hpMatches = [regex]::Matches($hpHTML, '(\w+ \d{1,2}, \d{4})[^<]*?Hotpatch\s+(KB\d+)[^<]*?OS Builds?\s+([\d.]+(?:\s+and\s+[\d.]+)*)')
    foreach ($m in $hpMatches) {
        $releaseD  = $m.Groups[1].Value.Trim()
        $kbNum     = $m.Groups[2].Value.Trim()
        $buildsRaw = $m.Groups[3].Value
        $buildNums = [regex]::Matches($buildsRaw, '(\d{4,5})\.(\d+)')
        foreach ($bm in $buildNums) {
            $majorBld = $bm.Groups[1].Value
            $minorBld = [int]$bm.Groups[2].Value
            $fullBld  = "$majorBld.$($bm.Groups[2].Value)"
            if ($majorBld -eq "26100" -or $majorBld -eq "26200") {
                if (-not $HotpatchOSBuilds.ContainsKey($majorBld) -or $minorBld -gt $HotpatchOSBuilds[$majorBld]) {
                    $HotpatchOSBuilds[$majorBld] = $minorBld
                    $HotpatchKBMap[$majorBld]    = $kbNum
                    $HotpatchDateMap[$majorBld]  = $releaseD
                }
                $osName = if ($majorBld -eq "26100") { "Win11 24H2" } else { "Win11 25H2" }
                $HotpatchAllBuilds += [PSCustomObject]@{
                    OperatingSystem = $osName
                    Build           = $fullBld
                    MajorBuild      = $majorBld
                    MinorBuild      = $bm.Groups[2].Value
                    PatchID         = $kbNum
                    ReleaseDate     = $releaseD
                    PatchTypeFlag   = "Hotpatch"
                    OSVersion       = "10.0.$fullBld"
                }
            }
        }
    }
    Write-Host "  Latest hotpatch thresholds:" -ForegroundColor Cyan
    foreach ($e in $HotpatchOSBuilds.GetEnumerator()) {
        Write-Host "    MajorBuild $($e.Key) -> MinorBuild $($e.Value)  ($($HotpatchKBMap[$e.Key]))" -ForegroundColor Cyan
    }
    Write-Host "  Total hotpatch build entries collected: $($HotpatchAllBuilds.Count)" -ForegroundColor Cyan
} catch {
    Write-Host "Warning: Could not scrape Hotpatch page ($_)" -ForegroundColor Yellow
}
Write-Host "Downloading and parsing Windows patch lists..." -ForegroundColor Yellow

# All active + historical Windows update history URLs
$PatchSourceURLs = @(
    "https://support.microsoft.com/en-us/topic/windows-11-version-25h2-update-history-99c7f493-df2a-4832-bd2d-6706baa0dec0",
    "https://support.microsoft.com/en-us/topic/windows-11-version-23h2-update-history-59875222-b990-4bd9-932f-91a5954de434",
    "https://support.microsoft.com/en-us/topic/windows-11-version-21h2-update-history-a19cd327-b57f-44b9-84e0-26ced7109ba9",
    "https://support.microsoft.com/en-us/topic/windows-10-update-history-8127c2c6-6edf-4fdf-8b9f-0f7be1ef3562"   # Win10 22H2 (and older Win10 builds)
)

$BuildDetails  = $buildInfoArray
$PatchDetails  = @()
$MajorBuilds   = @()

$entryRegex = '(\w+ \d{1,2}, \d{4})[^\n]*?(KB\d+)[^\n]*?OS Builds?\s+([\d.]+(?:\s*and\s*[\d.]+)*)'

foreach ($srcURL in $PatchSourceURLs) {
    Write-Host "  Fetching: $srcURL" -ForegroundColor Yellow
    try {
        $pageText = (Invoke-WebRequest -Uri $srcURL -UseBasicParsing -ErrorAction Stop).Content
        $matches2 = [regex]::Matches($pageText, $entryRegex)
        Write-Host "  Found $($matches2.Count) patch entries on this page" -ForegroundColor Cyan

        foreach ($m in $matches2) {
            $releaseDate = $m.Groups[1].Value.Trim()
            $patchKB     = $m.Groups[2].Value.Trim()
            $buildsRaw   = $m.Groups[3].Value
            $contextStart = [Math]::Max(0, $m.Index - 30)
            $contextLen   = [Math]::Min(120, $pageText.Length - $contextStart)
            $context      = $pageText.Substring($contextStart, $contextLen)
            $isPreview     = $context -match 'Preview'
            $isOOB         = $context -match 'out-of-band|Out-of-band|OOB'
            $buildNums = [regex]::Matches($buildsRaw, '(\d{4,5})\.(\d+)')
            foreach ($bm in $buildNums) {
                $fullBuild = $bm.Value.Trim()
                $mjBld     = $bm.Groups[1].Value
                $mnBld     = $bm.Groups[2].Value

                # Resolve OS name from buildInfoArray
                $osName = "Unknown"
                foreach ($bd in $BuildDetails) {
                    if ($bd.Build -eq $mjBld) { $osName = $bd.OperatingSystem; break }
                }

                $MajorBuilds += $mjBld

                # Build the patch type flag
                $patchTypeFlag = if ($isOOB) { "OOB" } elseif ($isPreview) { "Preview" } else { "PatchTuesday" }

                $PatchDetails += [PSCustomObject]@{
                    OperatingSystem = $osName
                    Build           = $fullBuild
                    MajorBuild      = $mjBld
                    MinorBuild      = $mnBld
                    PatchID         = $patchKB
                    ReleaseDate     = $releaseDate
                    PatchTypeFlag   = $patchTypeFlag
                }
            }
        }
    } catch {
        Write-Host "  Warning: Could not fetch $srcURL  ($_)" -ForegroundColor Yellow
    }
}

# Deduplicate and sort
$MajorBuilds  = $MajorBuilds | Select-Object -Unique | Sort-Object -Descending
$PatchDetails = $PatchDetails | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,PatchTypeFlag -Unique | Sort-Object MajorBuild,PatchID -Descending

Write-Host "Total patch entries loaded: $($PatchDetails.Count)" -ForegroundColor Green
Write-Host "Total unique major builds  : $($MajorBuilds.Count)" -ForegroundColor Green

# Build separate sets for Preview and OOB detection
$OSPreBuilds = ($PatchDetails | Where-Object { $_.PatchTypeFlag -eq "Preview" }  | Select-Object -ExpandProperty Build) | Sort-Object -Unique
$OSOOBBuilds = ($PatchDetails | Where-Object { $_.PatchTypeFlag -eq "OOB" }      | Select-Object -ExpandProperty Build) | Sort-Object -Unique

# PatchTuesday-only subset for export
$PTOnly = $PatchDetails | Where-Object { $_.PatchTypeFlag -eq "PatchTuesday" }
$PTOnly | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate | Export-Csv -Path $OutFileMP -NoTypeInformation

Write-Host "Patch Tuesday entries: $($PTOnly.Count)" -ForegroundColor Cyan
Write-Host "Preview entries      : $(($PatchDetails | Where-Object {$_.PatchTypeFlag -eq 'Preview'}).Count)" -ForegroundColor Cyan
Write-Host "OOB entries          : $(($PatchDetails | Where-Object {$_.PatchTypeFlag -eq 'OOB'}).Count)" -ForegroundColor Cyan
Write-Host "Selecting latest Patch Tuesday per major build..." -ForegroundColor Yellow
$LatestPatches = @()

IF ($PatchingMonth) {
    Foreach ($Bld in $MajorBuilds) {
        $LatestPatches += $PTOnly | Where-Object {
            $_.MajorBuild -eq $Bld -and
            $_.ReleaseDate -match $PatchingMonth.Year -and
            $_.ReleaseDate -match $PatchingMonth.Month
        } | Sort-Object {[int]$_.MinorBuild} -Descending | Select-Object -First 1
    }
} ELSE {
    $Today = Get-Date
    $LatestDate = ($PTOnly | Sort-Object {
        try { [datetime]::ParseExact($_.ReleaseDate, "MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
        catch { [datetime]::MinValue }
    } -Descending | Select-Object -First 1).ReleaseDate
    $DiffDays = ([datetime]$Today - [datetime]$LatestDate).Days

    IF ([int]$DiffDays -gt [int]$PatchReleaseDays) {
        Foreach ($Bld in $MajorBuilds) {
            $LatestPatches += $PTOnly | Where-Object { $_.MajorBuild -eq $Bld } |
                              Sort-Object {[int]$_.MinorBuild} -Descending | Select-Object -First 1
        }
    } ELSE {
        $Month = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -First 1
        $Year  = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -Last 1
        $PatchingMonth = [PSCustomObject]@{Month=$Month; Year=$Year}
        Foreach ($Bld in $MajorBuilds) {
            $LatestPatches += $PTOnly | Where-Object {
                $_.MajorBuild -eq $Bld -and
                $_.ReleaseDate -match $PatchingMonth.Year -and
                $_.ReleaseDate -match $PatchingMonth.Month
            } | Sort-Object {[int]$_.MinorBuild} -Descending | Select-Object -First 1
        }
        $M = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).Split(" ,") | Select-Object -First 1
        $Y = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).Split(" ,") | Select-Object -Last 1
        Foreach ($Bld1 in $MajorBuilds) {
            $Found = $LatestPatches | Where-Object { $_.MajorBuild -eq $Bld1 }
            IF (-not $Found) {
                $LatestPatches += $PTOnly | Where-Object {
                    $_.MajorBuild -eq $Bld1 -and $_.ReleaseDate -notlike "$M*$Y"
                } | Sort-Object {[int]$_.MinorBuild} -Descending | Select-Object -First 1
            }
        }
    }
}

$LatestPatches = $LatestPatches | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,
    @{Name="OSVersion"; Expression={"10.0.$($_.Build)"}}
$LatestPatches | Export-Csv -Path $OutFileLP -NoTypeInformation

$mostRecentDate = $LatestPatches | Sort-Object {
    try { [datetime]::ParseExact($_.ReleaseDate, "MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { [datetime]::MinValue }
} -Descending | Select-Object -First 1
$LatestDate = $mostRecentDate.ReleaseDate

Write-Host "Latest Patch Tuesday date: $LatestDate" -ForegroundColor Green
Write-Host "Latest patches loaded for $($LatestPatches.Count) major builds" -ForegroundColor Green

#==============================================================================================================================
# BUILD ALL-RELEASED-PATCHES MASTER LIST  (Patch Tuesday + Preview + OOB )
#==============================================================================================================================

$AllReleasedPatchs = $PatchDetails | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,
    @{Name="OSVersion"; Expression={"10.0.$($_.Build)"}}
$existingOSVersions = $AllReleasedPatchs | Select-Object -ExpandProperty OSVersion
foreach ($hp in $HotpatchAllBuilds) {
    if ($existingOSVersions -notcontains $hp.OSVersion) {
        $AllReleasedPatchs += $hp
        $existingOSVersions += $hp.OSVersion
    }
}
Write-Host "Total entries in AllReleasedPatchs after hotpatch merge: $($AllReleasedPatchs.Count)" -ForegroundColor Green

$AllReleasedPatchs | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,OSVersion |
    Export-Csv -Path $OutFileMP -NoTypeInformation

#==============================================================================================================================
# BUILD COMPLIANCE MASTER TABLE ($FinalReport)
#==============================================================================================================================

$IntuneDeviceHardwareInfo = @()
$HotpatchBuildDateLookup = @{}
foreach ($hpEntry in $HotpatchAllBuilds) {
    if (-not $HotpatchBuildDateLookup.ContainsKey($hpEntry.Build)) {
        try {
            $hpDt = [datetime]::ParseExact($hpEntry.ReleaseDate, "MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            $HotpatchBuildDateLookup[$hpEntry.Build] = $hpDt
        } catch { }
    }
}

# Pre-compute latest PT release date as [datetime] per major build - needed for hotpatch comparison
$LatestPTDatePerBuild = @{}
foreach ($lp in $LatestPatches) {
    try {
        $ptDt = [datetime]::ParseExact($lp.ReleaseDate, "MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
        $LatestPTDatePerBuild[$lp.MajorBuild] = $ptDt
    } catch { }
}

foreach ($AllReleasedPatch in $AllReleasedPatchs) {
    $deviceMinorBuild = [int]$AllReleasedPatch.MinorBuild
    $majorBld         = $AllReleasedPatch.MajorBuild
    $fullBuildKey     = $AllReleasedPatch.Build   # e.g. "26100.7985"

    $ptPatch  = $LatestPatches | Where-Object { $_.MajorBuild -eq $majorBld } | Select-Object -First 1
    $ptMinor  = if ($ptPatch) { [int]$ptPatch.MinorBuild } else { [int]::MaxValue }
    $ptDate   = if ($LatestPTDatePerBuild.ContainsKey($majorBld)) { $LatestPTDatePerBuild[$majorBld] } else { [datetime]::MaxValue }

    # Patch Tuesday compliance: device minor build >= latest PT minor build
    $isPTCompliant = ($ptMinor -ne [int]::MaxValue) -and ($deviceMinorBuild -ge $ptMinor)

    # Hotpatch compliance - TWO conditions must BOTH be true:
    # 1. Device's exact build must exist in $HotpatchAllBuilds
    # 2. That hotpatch's release date must be >= latest PT release date for this build family
    #    (the hotpatch is from the current or a more recent hotpatch cycle than the latest PT)
    $hpDate        = if ($HotpatchBuildDateLookup.ContainsKey($fullBuildKey)) { $HotpatchBuildDateLookup[$fullBuildKey] } else { $null }
    $isHPCompliant = ($null -ne $hpDate) -and ($hpDate -ge $ptDate)

    $isCompliant = $isPTCompliant -or $isHPCompliant
    $status      = if ($isCompliant) { "Compliant" } else { "Non-Compliant" }

    $timeSpan = try {
        (Get-Date).Subtract([DateTime]::ParseExact($AllReleasedPatch.ReleaseDate, "MMMM d, yyyy", [CultureInfo]::InvariantCulture))
    } catch { New-TimeSpan -Days 0 }

    $RequiredPatch = if ($isCompliant) { "Compliant" } else {
        $lp = $LatestPatches | Where-Object { $_.MajorBuild -eq $majorBld }
        if ($lp) { $lp.PatchID -join ", " } else { "BNE" }
    }
    $RequiredPatchRD = if ($isCompliant) { "Compliant" } else {
        $lp = $LatestPatches | Where-Object { $_.MajorBuild -eq $majorBld }
        if ($lp) { $lp.ReleaseDate -join ", " } else { "BNE" }
    }

    $IntuneDeviceHardwareInfo += [PSCustomObject][ordered]@{
        OperatingSystem = $AllReleasedPatch.OperatingSystem
        OSVersion       = $AllReleasedPatch.OSVersion
        Build           = $AllReleasedPatch.Build
        MajorBuild      = $majorBld
        MinorBuild      = $AllReleasedPatch.MinorBuild
        PatchID         = $AllReleasedPatch.PatchID
        ReleaseDate     = $AllReleasedPatch.ReleaseDate
        PatchStatus     = $status
        NotPatchSince   = if ($status -eq 'Compliant') { "Compliant" } else { $timeSpan.Days.ToString() + " days" }
        RequiredPatch   = $RequiredPatch
        RequiredPatchRD = $RequiredPatchRD
    }
}

$FinalReport = $IntuneDeviceHardwareInfo | Select-Object OperatingSystem,OSVersion,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,PatchStatus,NotPatchSince,RequiredPatch,RequiredPatchRD
$FinalReport | Export-Csv -Path $MergeOverallFile -NoTypeInformation

Write-Host "===============================Phase-2 (Downloading and Creating MS Patch List) ==========================================(Completed)" -ForegroundColor Green
Write-Host ""
Write-Host "===============================Phase-3 (Generating Windows Patching Compliance Report) =====================================(Started)" -ForegroundColor Green

$compliantCount          = 0
$manualCheckCount        = 0
$nonCompliantCount       = 0
$excl_JoinType           = 0
$excl_EOL                = 0
$excl_OutOfScope         = 0
$excl_Offline            = 0
$complianceReport        = @()
$totalDevices            = $DevicesInfos.Count
$progress                = 0

$OfflineCutoffDate = $null
try {
    $latestPTDateTime = [datetime]::ParseExact($LatestDate, "MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    $OfflineCutoffDate = $latestPTDateTime.AddDays($OfflineExclusionDays)
    Write-Host "Latest Patch Tuesday    : $LatestDate" -ForegroundColor Cyan
    Write-Host "Offline Cutoff Date     : $($OfflineCutoffDate.ToString('dd-MMM-yyyy')) (Patch Tuesday + $OfflineExclusionDays days)" -ForegroundColor Cyan
} catch {
    Write-Host "Warning: Could not parse LatestDate '$LatestDate' for offline cutoff calculation." -ForegroundColor Yellow
}
$WebTypeCache = @{}
foreach ($device in $DevicesInfos) {
    $deviceName      = $device."devicename"
    $Serialnumber    = $device."SerialNumber"
    $Manufacturer    = $device.Manufacturer
    $Model           = $device.Model
    $Managedby       = $device."Managed by"
    $SkuFamily       = $device.SkuFamily
    $Totalstorage    = ($device."Total storage" / 1024).ToString("N2")
    $Freestorage     = ($device."Free storage"  / 1024).ToString("N2")
    $deviceOSVersion = $device."osversion"                   # e.g. "10.0.26100.7979"
    $OSMajorBuild    = $deviceOSVersion.Split(".")[2]          # e.g. "26100"
    $OSMinorBuild    = $deviceOSVersion.Split(".")[3]          # e.g. "7979"

    $OSVersionV = if ($OSMajorBuild -eq "0" -or [string]::IsNullOrWhiteSpace($OSMajorBuild)) { "No OS version" } `
                  elseif ($OSMajorBuild -eq "7601") { "Win7-Or-Server" } `
                  elseif ($DynamicBuildMap.ContainsKey($OSMajorBuild)) { $DynamicBuildMap[$OSMajorBuild] } `
                  else { $deviceOSVersion }

    $Ownership    = $device.Ownership
    $Lastcheckin  = ""
    if ([string]::IsNullOrEmpty($device."Last check-in")) {
        $Lastcheckin = "No_CheckIndate"
    } else {
        $Lastcheckin = ([DateTime]::ParseExact($device."Last check-in", "yyyy-MM-dd HH:mm:ss.fffffff", $null)).ToString("dd-MMM-yy")
    }
    $Lastcheckin1       = $device."Last check-in"
    $LastcheckinDate    = [DateTime]::ParseExact($Lastcheckin1, "yyyy-MM-dd HH:mm:ss.fffffff", [System.Globalization.CultureInfo]::InvariantCulture)
    $today              = Get-Date -Format "yyyy-MM-dd"
    $Lastcheckin_Indays = (Get-Date $today) - $LastcheckinDate
    $deviceCategory     = ""
    if     ($Lastcheckin_Indays -eq "No_CheckIndate")                                          { $deviceCategory = "No date"       }
    elseif ($Lastcheckin_Indays.Days -ge 0  -and $Lastcheckin_Indays.Days -le 1)               { $deviceCategory = "0-1 days"      }
    elseif ($Lastcheckin_Indays.Days -ge 2  -and $Lastcheckin_Indays.Days -le 5)               { $deviceCategory = "2-5 days"      }
    elseif ($Lastcheckin_Indays.Days -ge 6  -and $Lastcheckin_Indays.Days -le 10)              { $deviceCategory = "5-10 days"     }
    elseif ($Lastcheckin_Indays.Days -ge 11 -and $Lastcheckin_Indays.Days -le 20)              { $deviceCategory = "11-20 days"    }
    elseif ($Lastcheckin_Indays.Days -ge 21 -and $Lastcheckin_Indays.Days -le 30)              { $deviceCategory = "21-30 days"    }
    elseif ($Lastcheckin_Indays.Days -ge 31 -and $Lastcheckin_Indays.Days -le 60)              { $deviceCategory = "31-60 days"    }
    elseif ($Lastcheckin_Indays.Days -ge 61 -and $Lastcheckin_Indays.Days -le 90)              { $deviceCategory = "61-90 days"    }
    elseif ($Lastcheckin_Indays.Days -gt 90)                                                   { $deviceCategory = "above 90 days" }
    #else                                           { $deviceCategory = "above 90 days" }
    else                                                                                       { $deviceCategory = "Check_Date"    }

    $JoinType       = $device.JoinType
    $PrimaryUserUPN = $device."Primary user UPN"

    # Compliance check: Patch Tuesday AND Hotpatch
    $deviceMinorInt = if ($OSMinorBuild -match '^\d+$') { [int]$OSMinorBuild } else { 0 }
    $deviceFullBuild = "$OSMajorBuild.$OSMinorBuild"   # e.g. "26100.7985"

    $ptPatch    = $LatestPatches | Where-Object { $_.MajorBuild -eq $OSMajorBuild } | Select-Object -First 1
    $ptMinorInt = if ($ptPatch) { [int]$ptPatch.MinorBuild } else { -1 }
    $ptDate     = if ($LatestPTDatePerBuild.ContainsKey($OSMajorBuild)) { $LatestPTDatePerBuild[$OSMajorBuild] } else { [datetime]::MaxValue }

    # Patch Tuesday compliance: device minor build >= latest PT minor build (same family)
    $isPTCompliant = ($ptMinorInt -ge 0) -and ($deviceMinorInt -ge $ptMinorInt)

    # Hotpatch compliance - BOTH conditions required:
    # 1. Exact full build (e.g. "26100.7985") must be in the hotpatch build list
    # 2. That hotpatch release date must be >= latest PT release date for this build family
    $hpBuildDate   = if ($HotpatchBuildDateLookup.ContainsKey($deviceFullBuild)) { $HotpatchBuildDateLookup[$deviceFullBuild] } else { $null }
    $isHotpatchBuild  = ($null -ne $hpBuildDate)
    $isHPCompliant = $isHotpatchBuild -and ($hpBuildDate -ge $ptDate)

    $complianceStatus = if ($isPTCompliant -or $isHPCompliant) { "Compliant" } else { "Non-Compliant" }

    # Lookup from FinalReport master table
    $frRow           = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -First 1
    $PatchID         = if ($frRow -and -not [string]::IsNullOrWhiteSpace($frRow.PatchID))        { $frRow.PatchID      } else { "Manually Check Installed KB"    }
    $ReleaseDate     = if ($frRow -and -not [string]::IsNullOrWhiteSpace($frRow.ReleaseDate))    { $frRow.ReleaseDate  } else { "Manually Check Release Date"     }
    $notPatchSince   = if ($frRow -and -not [string]::IsNullOrWhiteSpace($frRow.NotPatchSince))  { $frRow.NotPatchSince} else { "Manually Check"                  }
    $RequiredPatch   = if ($frRow -and -not [string]::IsNullOrWhiteSpace($frRow.RequiredPatch))  { $frRow.RequiredPatch} else { "Manually Check Required Patch"    }
    $RequiredPatchRD = if ($frRow -and -not [string]::IsNullOrWhiteSpace($frRow.RequiredPatchRD)){ $frRow.RequiredPatchRD } else { "Manually Check Required Patch" }

    
    if ($complianceStatus -eq "Compliant") {
        $notPatchSince   = "Compliant"
        $RequiredPatch   = "Compliant"
        $RequiredPatchRD = "Compliant"
    }
    function Get-PatchTypeFromWeb {
        param([string]$BuildNumber)   # e.g. "26100.8246"
        try {
            $searchUrl = "https://support.microsoft.com/en-us/search/results?query=OS+Build+$BuildNumber+Windows"
            $webContent = (Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -ErrorAction Stop -TimeoutSec 10).Content
            if ($webContent -match 'Hotpatch')           { return "Hotpatch Update"      }
            if ($webContent -match 'out-of-band|OOB')   { return "OOB Update"            }
            if ($webContent -match 'Preview')            { return "Preview Update"        }
            if ($webContent -match "OS Build $([regex]::Escape($BuildNumber))") {
                                                           return "Patch Tuesday Update"  }
        } catch { }
        
        try {
            $majorBldLocal = $BuildNumber.Split(".")[0]
            $historyUrls = @(
                "https://support.microsoft.com/en-us/topic/windows-11-version-25h2-update-history-99c7f493-df2a-4832-bd2d-6706baa0dec0",
                "https://support.microsoft.com/en-us/topic/windows-11-version-23h2-update-history-59875222-b990-4bd9-932f-91a5954de434",
                "https://support.microsoft.com/en-us/help/4043454"
            )
            foreach ($hUrl in $historyUrls) {
                $hContent = (Invoke-WebRequest -Uri $hUrl -UseBasicParsing -ErrorAction Stop -TimeoutSec 10).Content
                
                if ($hContent -match [regex]::Escape($BuildNumber)) {
                    $idx = ($hContent | Select-String -Pattern [regex]::Escape($BuildNumber)).Matches[0].Index
                    $ctx = $hContent.Substring([Math]::Max(0,$idx-80), [Math]::Min(200, $hContent.Length - [Math]::Max(0,$idx-80)))
                    if ($ctx -match 'Hotpatch')         { return "Hotpatch Update"      }
                    if ($ctx -match 'out-of-band|OOB')  { return "OOB Update"            }
                    if ($ctx -match 'Preview')          { return "Preview Update"        }
                    return "Patch Tuesday Update"
                }
            }
        } catch { }
        return $null   
    }

    $UpdateType = if ($complianceStatus -eq "Compliant") {
        if     ($isHotpatchBuild)                                                   { "Hotpatch Update"      }
        elseif ($ptMinorInt -ge 0 -and $deviceMinorInt -eq $ptMinorInt)             { "Patch Tuesday Update" }
        elseif ($deviceMinorInt -gt $ptMinorInt)                                    { "Expedited/OOB Update" }
        else                                                                        { "Patch Tuesday Update" }
    } elseif ($isHotpatchBuild)                          { "Hotpatch Update"      }
      elseif ($OSPreBuilds -contains $deviceFullBuild)   { "Preview Update"       }
      elseif ($OSOOBBuilds -contains $deviceFullBuild)   { "OOB Update"           }
     elseif ($PatchID -eq "Manually Check Installed KB") {
    if (-not $WebTypeCache.ContainsKey($deviceFullBuild)) {
        $WebTypeCache[$deviceFullBuild] = Get-PatchTypeFromWeb -BuildNumber $deviceFullBuild
    }
    if ($WebTypeCache[$deviceFullBuild]) { $WebTypeCache[$deviceFullBuild] } else { "Manually Check" }
}
      else {
          $localEntry = $AllReleasedPatchs | Where-Object { $_.Build -eq $deviceFullBuild } | Select-Object -First 1
          if     ($localEntry -and $localEntry.PatchTypeFlag -eq "Preview")      { "Preview Update"       }
          elseif ($localEntry -and $localEntry.PatchTypeFlag -eq "OOB")          { "OOB Update"            }
          else                                                                   { "Patch Tuesday Update" }
      }

    # EXCLUSION RULES 
    $exclusionReason = $null

    # Rule 1 - Join Type: only Hybrid Azure AD joined OR Azure AD joined are in scope
    if ($JoinType -ne "Hybrid Azure AD joined" -and $JoinType -ne "Azure AD joined") {
        $exclusionReason = "Exclusion (Join type not supported)"
    }

    # Rule 2 - Windows 10 EOL: any Win10 device is excluded
    if (-not $exclusionReason) {
        if ($OSVersionV -like "Win10*" -or $OSMajorBuild -in @("19045","19044","19043","19042","19041","18363","18362","17763","17134","16299","15063","14393","10586","10240")) {
            $exclusionReason = "Exclusion (EOL devices)"
        }
    }

    # Rule 3 - Out-of-scope device name prefixes
    if (-not $exclusionReason) {
        $outOfScopePrefixes = @($DevcieNamePrefix)
        foreach ($prefix in $outOfScopePrefixes) {
            if ($deviceName -like "$prefix*") {
                $exclusionReason = "Exclusion (Devices out of scope by name)"
                break
            }
        }
    }

    # Rule 4 - Offline during deployment: LastCheckin before (Patch Tuesday + X days)
   if (-not $exclusionReason -and $null -ne $OfflineCutoffDate -and $Lastcheckin -ne "No_CheckIndate" -and $complianceStatus -ne "Compliant") {
    if ((Get-Date) -gt $OfflineCutoffDate) {
        if ($LastcheckinDate -lt $OfflineCutoffDate) {
            $exclusionReason = "Exclusion (Devices offline from the day of patch deployment)"
        }
    }
}
    

    # Determine final PatchingComplianceStatus and update counters
    if ($exclusionReason) {
        $PatchingComplianceStatus = $exclusionReason
        # Increment the appropriate exclusion counter
        switch -Wildcard ($exclusionReason) {
            "*Join type not supported*"                    { $excl_JoinType++ }
            "*EOL devices*"                                { $excl_EOL++ }
            "*out of scope*"                               { $excl_OutOfScope++ }
            "*offline from the day of patch deployment*"   { $excl_Offline++ }
        }
    } else {
        # Device is in-scope - determine the correct PatchingComplianceStatus
    if ($complianceStatus -eq "Compliant") {
    $PatchingComplianceStatus = "Compliant"
    $compliantCount++
} elseif ($RequiredPatch -eq "Manually Check Required Patch" -and $OSMajorBuild -eq "0") {
    $PatchingComplianceStatus = "Manually Check Required Patch"
    $manualCheckCount++
} else {
    $PatchingComplianceStatus = "Non-Compliant"
    $nonCompliantCount++
} }

    $complianceReport += [PSCustomObject]@{
        "DeviceName"                 = $deviceName
        "Serialnumber"               = $Serialnumber
        "PrimaryUserUPN"             = $PrimaryUserUPN
        "Ownership"                  = $Ownership
        "JoinType"                   = $JoinType
        "Manufacturer"               = $Manufacturer
        "Model"                      = $Model
        "Managedby"                  = $Managedby
        "Totalstorage (GB)"          = $Totalstorage
        "Freestorage (GB)"           = $Freestorage
        "Lastcheckin"                = $Lastcheckin
        "Lastcheckin_Indays"         = $Lastcheckin_Indays.Days.ToString() + " days"
        "Lastcheckin_InBetween"      = $deviceCategory
        "SkuFamily"                  = $SkuFamily
        "OSVersion"                  = $OSVersionV
        "OS"                         = $deviceOSVersion
        "InstalledKB"                = $PatchID
        "UpdateType"                 = $UpdateType          
        "InstalledKB_ReleaseDate"    = $ReleaseDate
        "PatchingStatus"             = $complianceStatus
        "PatchingComplianceStatus"   = $PatchingComplianceStatus
        "DevcieNotPatchSince_InDays" = $notPatchSince
        "Latest_RequiredPatch"       = $RequiredPatch
        "RequiredPatchRD"            = $RequiredPatchRD
    }

    $progress++
    $percentComplete = [String]::Format("{0:0.00}", ($progress / $totalDevices) * 100)
    Write-Progress -Activity "Generating Windows Patching Compliance Report" -Status "Progress: $percentComplete% Complete" -PercentComplete $percentComplete
}

$complianceReport | Export-Csv -Path $Final_Patching_Report -NoTypeInformation

# In-scope total = Compliant + ManualCheck + NonCompliant (exclusions are NOT counted in compliance %)
$totalInScope         = $compliantCount + $manualCheckCount + $nonCompliantCount
$totalExcluded        = $excl_JoinType + $excl_EOL + $excl_OutOfScope + $excl_Offline
$totalAllDevices      = $totalInScope + $totalExcluded
$compliancePercentage = if ($totalInScope -gt 0) { "{0:N2}" -f ($compliantCount / $totalInScope * 100) } else { "0.00" }

Write-Host ""
Write-Host "=============================== Device Count Summary ============================================" -ForegroundColor Cyan
Write-Host "Total Devices (All)                                         : $totalAllDevices"
Write-Host "  >> In-Scope Devices (used for compliance %)               : $totalInScope"
Write-Host "  >> Total Excluded Devices                                 : $totalExcluded"
Write-Host ""
Write-Host "In-Scope Breakdown:"
Write-Host "  Compliant                                                 : $compliantCount"                             -ForegroundColor Green
Write-Host "  Non-Compliant                                             : $nonCompliantCount"                          -ForegroundColor Red
Write-Host "  Manually Check Required                                   : $manualCheckCount"                           -ForegroundColor Yellow
Write-Host ""
Write-Host "Excluded Breakdown:"
Write-Host "  Exclusion (Join type not supported)                       : $excl_JoinType"                              -ForegroundColor DarkGray
Write-Host "  Exclusion (EOL devices - Windows 10)                      : $excl_EOL"                                   -ForegroundColor DarkGray
Write-Host "  Exclusion (Devices out of scope by name)                  : $excl_OutOfScope"                            -ForegroundColor DarkGray
Write-Host "  Exclusion (Devices offline from patch deployment day)     : $excl_Offline"                               -ForegroundColor DarkGray
Write-Host ""
Write-Host "Patching Compliance Percentage (In-Scope Only)              : $compliancePercentage%" -ForegroundColor Green
Write-Host "================================================================================================" -ForegroundColor Cyan
Write-Host "===============================Phase-3 (Generating Windows Patching Compliance Report) ===================================(Completed)" -ForegroundColor Green

#==============================================================================================================================
# PHASE-4: CLEANUP
#==============================================================================================================================

Write-Host ""
Write-Host "===============================Phase-4 (Cleanup - Removing Intermediate Files) ===========================================(Started)" -ForegroundColor Green

$FilesToDelete = @(
    "$WorkingFolder\MergeOverallFile.csv",
    "$WorkingFolder\MicrosoftLatestPatchList.csv",
    "$WorkingFolder\MicrosoftPatchList.csv"
)

foreach ($file in $FilesToDelete) {
    if (Test-Path -Path $file) {
        Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
        Write-Host "  Deleted: $file" -ForegroundColor Yellow
    } else {
        Write-Host "  Not found (skipped): $file" -ForegroundColor DarkGray
    }
}

# Delete PD_Dump folder and all its contents
$PDDumpPath = "$WorkingFolder\PD_Dump"
if (Test-Path -Path $PDDumpPath) {
    Remove-Item -Path $PDDumpPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  Deleted folder: $PDDumpPath" -ForegroundColor Yellow
} else {
    Write-Host "  Not found (skipped): $PDDumpPath" -ForegroundColor DarkGray
}

# Also delete the downloaded zip file if it still exists in the working directory
$ZipFiles = Get-ChildItem -Path $WorkingFolder -Filter "*.zip" -ErrorAction SilentlyContinue
foreach ($zip in $ZipFiles) {
    Remove-Item -Path $zip.FullName -Force -ErrorAction SilentlyContinue
    Write-Host "  Deleted zip: $($zip.FullName)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Kept: $Final_Patching_Report" -ForegroundColor Green
#Invoke-item -Path "$WorkingFolder\Final_Patching_Report.csv"
Write-Host "===============================Phase-5 (Cleanup - Removing Intermediate Files) ===========================================(Completed)" -ForegroundColor Green
