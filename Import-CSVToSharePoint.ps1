<#
.SYNOPSIS
Ultra-fast CSV import into SharePoint Online using CSOM batching.

.DESCRIPTION
- Handles 10k–100k+ records
- Dynamic CSV → SharePoint mapping via ColumnMap
- Supports Choice & MultiChoice fields
- Optional Created/Modified preservation
- Throttling protection
- Test run support
#>

param(
    [Parameter(Mandatory=$true)][string]$SiteUrl,
    [Parameter(Mandatory=$true)][string]$ListName,
    [Parameter(Mandatory=$true)][string]$CsvPath,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [Parameter(Mandatory=$true)][string]$TenantName,

    [Parameter(Mandatory=$true)]
    [hashtable]$ColumnMap,   # "CSV Header" = "InternalName"

    [switch]$PreserveDates,
    [int]$BatchSize = 100,
    [int]$SleepAfterBatches = 10,
    [int]$SleepSeconds = 1,
    [switch]$TestRun,
    [int]$TestLimit = 100
)

Import-Module PnP.PowerShell -Force
Add-Type -AssemblyName "Microsoft.SharePoint.Client"

# ---------------- LOGGING ----------------
function Log {
    param($Message, $Type = "Info")
    $ts = Get-Date -Format "HH:mm:ss"
    $colors = @{ Info="Cyan"; Success="Green"; Warning="Yellow"; Error="Red" }
    Write-Host "[$ts] $Message" -ForegroundColor $colors[$Type]
}

# ---------------- CONNECT ----------------
function Ensure-PnPConnection {
    try {
        $connection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($connection -and $connection.Url -eq $SiteUrl) {
            Log "Already connected" "Success"
            return
        }
        Log "Connecting to SharePoint..."
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantName -DeviceLogin
        Log "Connected!" "Success"
    }
    catch {
        Write-Error "Connection failed: $($_.Exception.Message)"
        exit 1
    }
}

Ensure-PnPConnection

# ---------------- LOAD LIST ----------------
$ctx  = Get-PnPContext
$list = $ctx.Web.Lists.GetByTitle($ListName)
$ctx.Load($list)
$ctx.Load($list.Fields)
$ctx.ExecuteQuery()

# ---------------- DETECT CHOICE FIELDS ----------------
$ChoiceFields = @{}
$MultiChoiceFields = @{}

foreach ($f in $list.Fields) {
    if ($f.FieldTypeKind -eq [Microsoft.SharePoint.Client.FieldType]::Choice) {
        $ChoiceFields[$f.InternalName] = $f.Choices
    }
    elseif ($f.FieldTypeKind -eq [Microsoft.SharePoint.Client.FieldType]::MultiChoice) {
        $MultiChoiceFields[$f.InternalName] = $f.Choices
    }
}

# ---------------- DATE PARSER ----------------
function Parse-Date {
    param($Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try { return [DateTime]::Parse($Value) } catch { return $null }
}

# ---------------- MAP CSV ROW ----------------
function Map-CsvRow {
    param($Row)

    $mapped = @{}

    foreach ($csvHeader in $ColumnMap.Keys) {
        if (-not $Row.PSObject.Properties.Name.Contains($csvHeader)) { continue }

        $spField = $ColumnMap[$csvHeader]
        $value = $Row.$csvHeader

        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        # --- Single Choice ---
        if ($ChoiceFields.ContainsKey($spField)) {
            $valid = $ChoiceFields[$spField] | Where-Object { $_ -ieq $value }
            if ($valid) { $mapped[$spField] = $valid }
            else { Log "Invalid choice '$value' for '$spField'" "Warning" }
        }

        # --- Multi Choice ---
        elseif ($MultiChoiceFields.ContainsKey($spField)) {
            $choices = $MultiChoiceFields[$spField]
            $splitValues = $value -split '[,;]' | ForEach-Object { $_.Trim() }

            $validValues = @()
            foreach ($val in $splitValues) {
                $match = $choices | Where-Object { $_ -ieq $val }
                if ($match) { $validValues += $match }
                else { Log "Invalid multi-choice '$val' for '$spField'" "Warning" }
            }

            if ($validValues.Count -gt 0) {
                $mapped[$spField] = $validValues
            }
        }

        # --- Normal Field ---
        else {
            $mapped[$spField] = $value
        }
    }

    if ($PreserveDates) {
        $mapped["_Created"]  = $Row.Created
        $mapped["_Modified"] = $Row.Modified
    }

    return $mapped
}

# ---------------- BATCH CREATION ----------------
function Create-Batch {
    param($Data)

    $createdItems = @()

    foreach ($row in $Data) {
        $info = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $item = $list.AddItem($info)

        foreach ($key in $row.Keys) {
            if ($key -notlike "_*") { $item[$key] = $row[$key] }
        }

        if ([string]::IsNullOrWhiteSpace($item["Title"])) {
            $item["Title"] = "Imported_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        }

        $item.Update()
        $createdItems += @{ Item = $item; Meta = $row }
    }

    $ctx.ExecuteQuery()

    if ($PreserveDates) {
        foreach ($entry in $createdItems) {
            $needsUpdate = $false
            $c = Parse-Date $entry.Meta["_Created"]
            $m = Parse-Date $entry.Meta["_Modified"]

            if ($c) { $entry.Item["Created"] = $c; $needsUpdate = $true }
            if ($m) { $entry.Item["Modified"] = $m; $needsUpdate = $true }

            if ($needsUpdate) { $entry.Item.UpdateOverwriteVersion() }
        }
        $ctx.ExecuteQuery()
    }

    return $createdItems.Count
}

# ---------------- MAIN ----------------
if (!(Test-Path $CsvPath)) { Write-Error "CSV not found"; exit 1 }

$csv = Import-Csv $CsvPath
$totalRecords = $csv.Count

if ($TestRun) {
    $csv = $csv | Select-Object -First $TestLimit
    $totalRecords = $csv.Count
    Log "Test run: $totalRecords records" "Warning"
}

Log "Starting import: $totalRecords items"

$processed = 0
$batchNum = 1
$startTime = Get-Date

for ($i=0; $i -lt $totalRecords; $i+=$BatchSize) {
    $batchRows = $csv[$i..([Math]::Min($i+$BatchSize-1,$totalRecords-1))]
    $batchData = $batchRows | ForEach-Object { Map-CsvRow $_ }

    Log "Batch $batchNum started ($($batchData.Count) items)"
    $count = Create-Batch $batchData

    $processed += $count
    Log "Batch $batchNum completed"

    if ($batchNum % $SleepAfterBatches -eq 0) { Start-Sleep -Seconds $SleepSeconds }
    $batchNum++
}

$duration = (Get-Date) - $startTime
Log "Import finished: $processed items in $($duration.ToString('hh\:mm\:ss'))" "Success"
