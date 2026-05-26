# build_phase1.ps1 - ASCII only
# Builds v2 skeleton xlsx (17 worksheets + TODAY centralization + freeze panes)
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build_phase1.ps1

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Paths (resolve relative to script location)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildDir  = Split-Path -Parent $scriptDir
$configPath = Join-Path $buildDir 'config\phase1_skeleton.json'
$outputDir  = Join-Path $buildDir 'output'
$logDir     = Join-Path $buildDir 'log'
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $logDir "build_phase1_$ts.log"

# Output file name from config (read later); set placeholder
$outputName = 'PoliceStation_v2.0_skeleton.xlsx'
$outputPath = Join-Path $outputDir $outputName

# Logger
function Write-Log {
    param([string]$msg, [string]$level = 'INFO')
    $line = "$(Get-Date -Format 'HH:mm:ss') [$level] $msg"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

# Start
"" | Out-File -LiteralPath $logPath -Encoding UTF8
Write-Log "===== Phase 1 build started ====="
Write-Log "Script dir: $scriptDir"
Write-Log "Build dir: $buildDir"
Write-Log "Config: $configPath"

# Load config (UTF-8)
if (-not (Test-Path -LiteralPath $configPath)) {
    Write-Log "Config not found: $configPath" 'ERROR'
    exit 1
}

try {
    $configRaw = Get-Content -LiteralPath $configPath -Encoding UTF8 -Raw
    $config = $configRaw | ConvertFrom-Json
    Write-Log "Config loaded: version $($config.version), $($config.sheets.Count) sheets"
} catch {
    Write-Log "Failed to parse JSON: $_" 'ERROR'
    exit 1
}

# Backup existing output if exists
if (Test-Path -LiteralPath $outputPath) {
    $bakPath = $outputPath -replace '\.xlsx$', "_bak_$ts.xlsx"
    Move-Item -LiteralPath $outputPath -Destination $bakPath -Force
    Write-Log "Backed up existing output to $bakPath"
}

# Start Excel COM
Write-Log "Starting Excel COM..."
$excel = $null
$wb = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    Write-Log ("Excel started, version " + $excel.Version)

    $wb = $excel.Workbooks.Add()
    Write-Log ("Workbook created, default sheet count: " + $wb.Worksheets.Count)

    # Reduce to 1 sheet
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    # Sort sheets by order
    $sheetsSorted = $config.sheets | Sort-Object -Property order

    # Track sheet refs in hashtable (avoids Worksheets.Item lookup with Chinese names later)
    $sheetMap = @{}

    # Create / rename worksheets
    for ($i = 0; $i -lt $sheetsSorted.Count; $i++) {
        $s = $sheetsSorted[$i]
        if ($i -eq 0) {
            $ws = $wb.Worksheets.Item(1)
        } else {
            # Add new sheet after the last one
            $lastWs = $wb.Worksheets.Item($wb.Worksheets.Count)
            $ws = $wb.Worksheets.Add([Type]::Missing, $lastWs)
        }
        $ws.Name = $s.name
        $sheetMap[$s.name] = $ws
        Write-Log ("Sheet [{0}] {1} (freeze: {2})" -f $s.order, $s.name, $s.freeze)

        # Marker cell A1 (text from config template)
        $marker = $config.skeleton_marker.text_template.Replace('{sheet_name}', $s.name)
        $markerCell = $config.skeleton_marker.cell
        $ws.Range($markerCell).Value2 = $marker
        $ws.Range($markerCell).Font.Bold = $true
        $ws.Range($markerCell).Font.Size = 14
        $ws.Range($markerCell).Font.Color = 0x794E1F  # navy in BGR

        # Default font for sheet
        $ws.Cells.Font.Name = $config.default_font.family
        $ws.Cells.Font.Size = $config.default_font.size_body
    }

    # Freeze panes (use sheetMap to avoid Chinese name lookup)
    foreach ($s in $sheetsSorted) {
        if ($s.freeze) {
            $ws = $sheetMap[$s.name]
            $ws.Activate()
            $ws.Range($s.freeze).Select()
            $excel.ActiveWindow.FreezePanes = $true
            Write-Log ("Freeze pane set: {0} at {1}" -f $s.name, $s.freeze)
        }
    }

    # TODAY centralization on 設定表 (use sheetMap)
    $todayCfg = $config.today_centralization
    $settingsWs = $sheetMap[$todayCfg.sheet]
    if (-not $settingsWs) {
        throw "Settings sheet not found in sheetMap: $($todayCfg.sheet) (keys: $($sheetMap.Keys -join ','))"
    }
    $settingsWs.Range($todayCfg.label_cell).Value2 = $todayCfg.label_text
    $settingsWs.Range($todayCfg.label_cell).Font.Bold = $true
    $settingsWs.Range($todayCfg.formula_cell).Formula = $todayCfg.formula
    $settingsWs.Range($todayCfg.formula_cell).NumberFormat = 'yyyy/mm/dd'
    $settingsWs.Range($todayCfg.formula_cell).Font.Bold = $true
    $settingsWs.Range($todayCfg.formula_cell).Interior.Color = 0xE1F8FF  # BGR for #FFF8E1
    Write-Log ("TODAY centralized at {0}!{1}" -f $todayCfg.sheet, $todayCfg.formula_cell)

    # Named ranges - use full-signature call with named parameters to avoid COM dispatch issues
    foreach ($nr in $config.named_ranges_phase1) {
        try {
            # Names.Add signature: Name, RefersTo, Visible, MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal, RefersToR1C1, RefersToR1C1Local
            $null = $wb.Names.Add([string]$nr.name, [string]$nr.refers_to, $true)
            Write-Log ("Named range added: {0} = {1}" -f $nr.name, $nr.refers_to)
        } catch {
            Write-Log ("Failed to add named range {0}: {1}" -f $nr.name, $_.Exception.Message) 'WARN'
        }
    }

    # Activate first sheet for save (by index, not Chinese name)
    $wb.Worksheets.Item(1).Activate()

    # Save as xlsx
    Write-Log "Saving to $outputPath"
    $wb.SaveAs($outputPath, [int]$config.file_format_code)
    Write-Log ("File saved, size: " + (Get-Item -LiteralPath $outputPath).Length + " bytes")

} catch {
    Write-Log "EXCEPTION: $($_.Exception.Message)" 'ERROR'
    Write-Log "Stack: $($_.ScriptStackTrace)" 'ERROR'
    $exitCode = 1
} finally {
    if ($wb) {
        try { $wb.Close($false) | Out-Null } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    Write-Log "===== Phase 1 build ended (exit $exitCode) ====="
}

exit $exitCode
