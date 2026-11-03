<#
.SYNOPSIS
    Exports all Azure Update Manager (AUM) maintenance schedules (InGuestPatch maintenance configurations)
    with recurrence in human-readable form, next run date, window end, reboot setting, classifications
    and number of assigned machines / dynamic scopes.

.WHERE TO RUN
    Azure Cloud Shell (PowerShell)  -or-  a workstation with PowerShell 5.1+ and the Az.Accounts + Az.ResourceGraph modules.
    (Install-Module Az.ResourceGraph -Scope CurrentUser   if it's missing locally)

.OUTPUT
    AUM_Schedules_<yyyyMMdd_HHmm>.csv in the current folder (Cloud Shell: $HOME/clouddrive is persistent).
#>

[CmdletBinding()]
param(
    [string]$TenantId = '',            # if empty, the script prompts for it
    [string]$NameFilter = '',          # e.g. 'PRD-NAMR' -> only schedules whose name contains this text
    [string]$OutputFolder = (Get-Location).Path
)

#region --- Prerequisites / tenant login --------------------------------------------
Import-Module Az.Accounts      -ErrorAction Stop
Import-Module Az.ResourceGraph -ErrorAction Stop

while ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
    if ($TenantId) { Write-Host "Invalid TenantId format, try again." -ForegroundColor Yellow }
    $TenantId = (Read-Host "Enter the Tenant ID to connect to").Trim()
}

# Reuse the session only if it is already on the requested tenant; otherwise log in to that tenant
$ctx = Get-AzContext -ErrorAction SilentlyContinue
if (-not $ctx -or $ctx.Tenant.Id -ne $TenantId) {
    Write-Host "Connecting to tenant $TenantId ..." -ForegroundColor Cyan
    Connect-AzAccount -TenantId $TenantId -ErrorAction Stop | Out-Null
    $ctx = Get-AzContext
}
Write-Host ("Connected as {0} | Tenant: {1}" -f $ctx.Account.Id, $ctx.Tenant.Id) -ForegroundColor Green
#endregion

#region --- Resource Graph query ---------------------------------------------------
$query = @'
resources
| where type =~ 'microsoft.maintenance/maintenanceconfigurations'
| where tostring(properties.maintenanceScope) =~ 'InGuestPatch'
| extend mw = properties.maintenanceWindow, ip = properties.installPatches
| project
    Id               = tolower(id),
    ScheduleName     = name,
    SubscriptionId   = subscriptionId,
    ResourceGroup    = resourceGroup,
    Location         = location,
    StartDateTime    = tostring(mw.startDateTime),
    ExpirationDate   = tostring(mw.expirationDateTime),
    Duration         = tostring(mw.duration),
    TimeZone         = tostring(mw.timeZone),
    RecurEvery       = tostring(mw.recurEvery),
    RebootSetting    = tostring(ip.rebootSetting),
    WinClassifications = strcat_array(ip.windowsParameters.classificationsToInclude, ', '),
    WinKBExclude     = strcat_array(ip.windowsParameters.kbNumbersToExclude, ', '),
    WinKBInclude     = strcat_array(ip.windowsParameters.kbNumbersToInclude, ', '),
    LinuxClassifications = strcat_array(ip.linuxParameters.classificationsToInclude, ', ')
| join kind=leftouter (
    maintenanceresources
    | where type =~ 'microsoft.maintenance/configurationassignments'
    | extend cfg = tolower(tostring(properties.maintenanceConfigurationId)),
             isDynamic = isnotnull(properties['filter'])
    | summarize StaticMachines = countif(isDynamic == false),
                DynamicScopes  = countif(isDynamic == true) by cfg
  ) on $left.Id == $right.cfg
| project-away cfg
| order by ScheduleName asc
'@

$results = @()
$skip = $null
do {
    $params = @{ Query = $query; First = 1000; UseTenantScope = $true }
    if ($skip) { $params.SkipToken = $skip }
    $page = Search-AzGraph @params
    $results += $page.Data
    $skip = $page.SkipToken
} while ($skip)

if ($NameFilter) { $results = $results | Where-Object { $_.ScheduleName -like "*$NameFilter*" } }
Write-Host ("Schedules found: {0}" -f $results.Count) -ForegroundColor Cyan
#endregion

#region --- Helpers -----------------------------------------------------------------
$ordinalMap = @{ 'First' = 1; 'Second' = 2; 'Third' = 3; 'Fourth' = 4; 'Last' = -1 }

function Get-NthWeekday {
    param([int]$Year, [int]$Month, [int]$Ordinal, [System.DayOfWeek]$Day)
    if ($Ordinal -eq -1) {
        $d = (New-Object DateTime $Year, $Month, 1).AddMonths(1).AddDays(-1)
        while ($d.DayOfWeek -ne $Day) { $d = $d.AddDays(-1) }
        return $d
    }
    $d = New-Object DateTime $Year, $Month, 1
    while ($d.DayOfWeek -ne $Day) { $d = $d.AddDays(1) }
    return $d.AddDays(7 * ($Ordinal - 1))
}

function Get-NowInZone {
    param([string]$TimeZoneId)
    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    } catch { return [DateTime]::UtcNow }
}

function ConvertFrom-RecurEvery {
    # Returns: Friendly text + next run (in the schedule's own time zone)
    param([string]$Recur, [string]$Start, [string]$TimeZoneId)

    $startDt  = [DateTime]::Parse($Start)
    $now      = Get-NowInZone -TimeZoneId $TimeZoneId
    $from     = if ($startDt -gt $now) { $startDt } else { $now }
    $friendly = $Recur
    $next     = $null

    # Monthly - e.g. "Month Second Tuesday Offset1" / "1Month Last Sunday"
    if ($Recur -match '^\s*(\d*)\s*Month\s+(First|Second|Third|Fourth|Last)\s+(\w+day)(?:\s+Offset(-?\d+))?\s*$') {
        $interval = if ($Matches[1]) { [int]$Matches[1] } else { 1 }
        $ordName  = $Matches[2]
        $dayName  = $Matches[3]
        $offset   = if ($Matches[4]) { [int]$Matches[4] } else { 0 }

        if     ($offset -gt 0) { $prefix = "$offset day(s) after the" }
        elseif ($offset -lt 0) { $prefix = "$([math]::Abs($offset)) day(s) before the" }
        else                   { $prefix = 'The' }
        $friendly = "$prefix $ordName $dayName every $interval Month(s)"

        $dow = [System.DayOfWeek]$dayName
        $cursor = New-Object DateTime $from.Year, $from.Month, 1
        for ($i = 0; $i -lt 26; $i++) {
            $m = $cursor.AddMonths($i)
            # honour interval counting from the start month
            $monthsFromStart = (($m.Year - $startDt.Year) * 12) + ($m.Month - $startDt.Month)
            if ($monthsFromStart % $interval -ne 0) { continue }
            $cand = (Get-NthWeekday -Year $m.Year -Month $m.Month -Ordinal $ordinalMap[$ordName] -Day $dow).AddDays($offset)
            $cand = $cand.Date.Add($startDt.TimeOfDay)
            if ($cand -ge $from) { $next = $cand; break }
        }
    }
    # Monthly by day number - e.g. "Month day15,day30"
    elseif ($Recur -match '^\s*(\d*)\s*Month\s+((day\d+,?\s*)+)$') {
        $days = [regex]::Matches($Matches[2], 'day(\d+)') | ForEach-Object { [int]$_.Groups[1].Value }
        $friendly = "Monthly on day(s) " + ($days -join ', ')
        for ($i = 0; $i -lt 400 -and -not $next; $i++) {
            $d = $from.Date.AddDays($i).Add($startDt.TimeOfDay)
            if ($days -contains $d.Day -and $d -ge $from) { $next = $d }
        }
    }
    # Weekly - e.g. "Week Saturday,Sunday" / "2Week Friday"
    elseif ($Recur -match '^\s*(\d*)\s*Week\s+(.+)$') {
        $interval = if ($Matches[1]) { [int]$Matches[1] } else { 1 }
        $days = $Matches[2] -split '\s*,\s*'
        $friendly = "Every $interval Week(s) on " + ($days -join ', ')
        for ($i = 0; $i -lt 15 -and -not $next; $i++) {
            $d = $from.Date.AddDays($i).Add($startDt.TimeOfDay)
            if ($days -contains $d.DayOfWeek.ToString() -and $d -ge $from) { $next = $d }
        }
        if ($interval -gt 1) { $friendly += ' (next run approx. for interval > 1)' }
    }
    # Daily - e.g. "Day" / "3Day"
    elseif ($Recur -match '^\s*(\d*)\s*Day\s*$') {
        $interval = if ($Matches[1]) { [int]$Matches[1] } else { 1 }
        $friendly = "Every $interval Day(s)"
        $d = $from.Date.Add($startDt.TimeOfDay)
        if ($d -lt $from) { $d = $d.AddDays(1) }
        $next = $d
    }

    [pscustomobject]@{ Friendly = $friendly; Next = $next }
}
#endregion

#region --- Build output ------------------------------------------------------------
$report = foreach ($r in $results) {
    $parsed = ConvertFrom-RecurEvery -Recur $r.RecurEvery -Start $r.StartDateTime -TimeZoneId $r.TimeZone

    $windowEnd = $null
    if ($parsed.Next -and $r.Duration -match '^(\d{1,2}):(\d{2})$') {
        $windowEnd = $parsed.Next.AddHours([int]$Matches[1]).AddMinutes([int]$Matches[2])
    }

    [pscustomobject][ordered]@{
        'Schedule Name'          = $r.ScheduleName
        'Repeats Every'          = $parsed.Friendly
        'RecurEvery (raw)'       = $r.RecurEvery
        'Start Time'             = ([DateTime]::Parse($r.StartDateTime)).ToString('HH:mm')
        'Duration (HH:mm)'       = $r.Duration
        'Time Zone'              = $r.TimeZone
        'Next Run (local TZ)'    = if ($parsed.Next) { $parsed.Next.ToString('dd/MM/yyyy HH:mm') } else { 'N/A' }
        'Window End (local TZ)'  = if ($windowEnd)   { $windowEnd.ToString('dd/MM/yyyy HH:mm') }   else { 'N/A' }
        'Reboot Setting'         = $r.RebootSetting
        'Windows Classifications'= $r.WinClassifications
        'KB Exclusions'          = $r.WinKBExclude
        'KB Inclusions'          = $r.WinKBInclude
        'Linux Classifications'  = $r.LinuxClassifications
        'Static Machines'        = if ($r.StaticMachines) { $r.StaticMachines } else { 0 }
        'Dynamic Scopes'         = if ($r.DynamicScopes)  { $r.DynamicScopes }  else { 0 }
        'Schedule Start Date'    = ([DateTime]::Parse($r.StartDateTime)).ToString('dd/MM/yyyy')
        'Expiration'             = $r.ExpirationDate
        'Subscription Id'        = $r.SubscriptionId
        'Resource Group'         = $r.ResourceGroup
        'Location'               = $r.Location
        'Resource Id'            = $r.Id
    }
}

$outFile = Join-Path $OutputFolder ("AUM_Schedules_{0}_{1}.csv" -f $TenantId.Substring(0,8), (Get-Date -Format 'yyyyMMdd_HHmm'))
$report | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

$report | Sort-Object 'Schedule Name' |
    Format-Table 'Schedule Name', 'Repeats Every', 'Start Time', 'Next Run (local TZ)', 'Static Machines', 'Dynamic Scopes' -AutoSize

Write-Host "Exported to: $outFile" -ForegroundColor Green
#endregion