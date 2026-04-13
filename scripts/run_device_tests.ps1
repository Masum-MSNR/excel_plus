<#
.SYNOPSIS
  Boots an Android emulator, runs excel_plus integration tests, reports results.

.PARAMETER Avd
  AVD name to launch. Default: Small_Phone

.PARAMETER Cold
  If set, cold-boots the emulator (no snapshot). Default: uses snapshot.

.EXAMPLE
  .\run_device_tests.ps1
  .\run_device_tests.ps1 -Avd Pixel_Tablet
  .\run_device_tests.ps1 -Avd Small_Phone -Cold
#>
param(
    [string]$Avd = "Small_Phone",
    [switch]$Cold
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$exampleDir = Join-Path $root "example"

# --- Resolve Android SDK ---
$sdk = if ($env:ANDROID_HOME) { $env:ANDROID_HOME }
       elseif ($env:ANDROID_SDK_ROOT) { $env:ANDROID_SDK_ROOT }
       else { Join-Path $env:LOCALAPPDATA "Android\Sdk" }
$emulatorExe = Join-Path $sdk "emulator\emulator.exe"
$adbExe = Join-Path $sdk "platform-tools\adb.exe"

if (-not (Test-Path $emulatorExe)) {
    Write-Error "emulator.exe not found at $emulatorExe - set ANDROID_HOME"
    exit 1
}

# --- Check if emulator is already running ---
$existingDevice = & $adbExe devices 2>&1 | Select-String "emulator-\d+" | Select-Object -First 1
$emulatorStarted = $false

if ($existingDevice) {
    $deviceId = ($existingDevice -split "\s+")[0]
    Write-Host ":: Emulator already running: $deviceId" -ForegroundColor Green
} else {
    # --- Launch emulator ---
    Write-Host ":: Launching emulator '$Avd'..." -ForegroundColor Cyan
    $emulatorArgs = @("-avd", $Avd, "-no-audio", "-no-window")
    if ($Cold) { $emulatorArgs += "-no-snapshot-load" }

    $emulatorProcess = Start-Process -FilePath $emulatorExe -ArgumentList $emulatorArgs -PassThru -WindowStyle Hidden
    $emulatorStarted = $true

    # Wait for device to appear
    Write-Host ":: Waiting for device to come online..." -ForegroundColor Cyan
    $timeout = 120
    $elapsed = 0
    while ($elapsed -lt $timeout) {
        $dev = & $adbExe devices 2>&1 | Select-String "emulator-\d+\s+device"
        if ($dev) { break }
        Start-Sleep -Seconds 2
        $elapsed += 2
        Write-Host "   $($elapsed)s" -NoNewline
    }
    Write-Host ""

    if ($elapsed -ge $timeout) {
        Write-Error "Emulator did not come online within $($timeout)s"
        Stop-Process -Id $emulatorProcess.Id -Force -ErrorAction SilentlyContinue
        exit 1
    }

    $deviceId = ((& $adbExe devices 2>&1 | Select-String "emulator-\d+") -split "\s+")[0]
    Write-Host ":: Device online: $deviceId" -ForegroundColor Green

    # Wait for boot_completed
    Write-Host ":: Waiting for boot to complete..." -ForegroundColor Cyan
    $bootTimeout = 120
    $bootElapsed = 0
    while ($bootElapsed -lt $bootTimeout) {
        $bootProp = & $adbExe -s $deviceId shell getprop sys.boot_completed 2>&1
        if ($bootProp.Trim() -eq "1") { break }
        Start-Sleep -Seconds 2
        $bootElapsed += 2
    }

    if ($bootElapsed -ge $bootTimeout) {
        Write-Error "Emulator boot did not complete within $($bootTimeout)s"
        exit 1
    }
    Write-Host ":: Boot complete!" -ForegroundColor Green
}

# --- Run integration tests ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Running integration tests on $deviceId" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

Push-Location $exampleDir
try {
    & flutter test integration_test/excel_test.dart -d $deviceId --no-pub 2>&1 | Tee-Object -Variable testOutput
    $testExitCode = $LASTEXITCODE
} finally {
    Pop-Location
}

# --- Report ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
if ($testExitCode -eq 0) {
    Write-Host "  ALL TESTS PASSED" -ForegroundColor Green
} else {
    Write-Host "  TESTS FAILED (exit code: $testExitCode)" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Yellow

# --- Optionally shut down emulator we started ---
if ($emulatorStarted) {
    Write-Host ""
    Write-Host ":: Shutting down emulator..." -ForegroundColor Cyan
    & $adbExe -s $deviceId emu kill 2>&1 | Out-Null
}

exit $testExitCode
