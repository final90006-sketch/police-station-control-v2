# verify.ps1 - ASCII only
# Opens the built xlsx and verifies structure
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\verify.ps1

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildDir  = Split-Path -Parent $scriptDir
$configPath = Join-Path $buildDir 'config\phase1_skeleton.json'
$outputDir  = Join-Path $buildDir 'output'
$logDir     = Join-Path $buildDir 'log'
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $logDir "verify_$ts.log"

$outputName = 'PoliceStation_v2.0_skeleton.xlsx'
$outputPath = Join-Path $outputDir $outputName

function Write-Log {
    param([string]$msg, [string]$level = 'INFO')
    $line = "$(Get-Date -Format 'HH:mm:ss') [$level] $msg"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

"" | Out-File -LiteralPath $logPath -Encoding UTF8
Write-Log "===== Verify started ====="
Write-Log "Target: $outputPath"

$passCount = 0
$failCount = 0

function Test-Item {
    param([string]$name, [scriptblock]$check)
    try {
        $result = & $check
        if ($result) {
            Write-Log "[PASS] $name"
            $script:passCount++
            return $true
        } else {
            Write-Log "[FAIL] $name" 'WARN'
            $script:failCount++
            return $false
        }
    } catch {
        Write-Log "[FAIL] $name -- exception: $($_.Exception.Message)" 'ERROR'
        $script:failCount++
        return $false
    }
}

# 1. File exists
if (-not (Test-Path -LiteralPath $outputPath)) {
    Write-Log "Output file not found: $outputPath" 'ERROR'
    exit 1
}
Write-Log ("File size: " + (Get-Item -LiteralPath $outputPath).Length + " bytes")

# Load config
$config = Get-Content -LiteralPath $configPath -Encoding UTF8 -Raw | ConvertFrom-Json

# Start Excel COM
$excel = $null
$wb = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Log "Opening workbook..."
    $wb = $excel.Workbooks.Open($outputPath, 0, $true)  # ReadOnly = true
    Write-Log ("Workbook opened, sheets: " + $wb.Worksheets.Count)

    # === Tests ===

    # Test 1: 17 worksheets
    Test-Item "Worksheet count = 17" { $wb.Worksheets.Count -eq 17 }

    # Test 2: Sheet names + order
    $expectedSheets = $config.sheets | Sort-Object -Property order
    for ($i = 0; $i -lt $expectedSheets.Count; $i++) {
        $expected = $expectedSheets[$i]
        $actual = $wb.Worksheets.Item($i + 1).Name
        Test-Item ("Sheet[" + ($i+1) + "] name = " + $expected.name) { $actual -eq $expected.name }
    }

    # Test 3: TODAY centralization
    $todayCfg = $config.today_centralization
    $todayWs = $wb.Worksheets.Item($todayCfg.sheet)
    $todayCell = $todayWs.Range($todayCfg.formula_cell)
    Test-Item "TODAY formula at $($todayCfg.sheet)!$($todayCfg.formula_cell)" {
        $todayCell.Formula -eq $todayCfg.formula
    }
    Test-Item "TODAY formula evaluates to today" {
        $val = $todayCell.Value2
        $today = [Math]::Floor((Get-Date).ToOADate())
        [Math]::Floor($val) -eq $today
    }

    # Test 4: Named range
    foreach ($nr in $config.named_ranges_phase1) {
        Test-Item "Named range '$($nr.name)' exists" {
            $exists = $false
            foreach ($n in $wb.Names) {
                if ($n.Name -eq $nr.name) { $exists = $true; break }
            }
            $exists
        }
    }

    # Test 5: Freeze panes (sample-check a few)
    foreach ($s in @($expectedSheets[0], $expectedSheets[5], $expectedSheets[14])) {
        if ($s.freeze) {
            $ws = $wb.Worksheets.Item($s.name)
            $ws.Activate()
            Test-Item "Freeze pane on $($s.name)" {
                $excel.ActiveWindow.FreezePanes -eq $true
            }
        }
    }

    # Test 6: No #REF / #NAME errors on A1 markers (sample)
    foreach ($s in $expectedSheets) {
        $ws = $wb.Worksheets.Item($s.name)
        $v = $ws.Range("A1").Value2
        $ok = ($v -is [string]) -and ($v.Length -gt 0) -and (-not $v.StartsWith('#'))
        if (-not $ok) {
            Write-Log "[WARN] $($s.name) A1 marker invalid: $v" 'WARN'
        }
    }

} catch {
    Write-Log "EXCEPTION: $($_.Exception.Message)" 'ERROR'
    $exitCode = 1
} finally {
    if ($wb) {
        try { $wb.Close($false) } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

Write-Log "===== Verify ended ====="
Write-Log ("Total: $($passCount + $failCount) | Pass: $passCount | Fail: $failCount")

if ($failCount -gt 0) {
    Write-Log "FAILURES PRESENT - DO NOT PROCEED" 'ERROR'
    exit 1
}
exit $exitCode
