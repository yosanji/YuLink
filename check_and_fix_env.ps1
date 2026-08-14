# YuLink - Environment Check & Auto-Repair Tool
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "       YuLink (YuLian) - Environment Check & Auto-Repair Tool    " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check .NET Framework 4.8+
Write-Host "[1/3] Checking .NET Framework 4.8+ runtime..." -ForegroundColor Yellow
$net48Installed = $false
try {
    $release = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
    if ($release -and $release -ge 528040) {
        $net48Installed = $true
    }
} catch {
    $net48Installed = $false
}

if ($net48Installed) {
    Write-Host "  --> [OK] .NET Framework 4.8+ detected (Release: $release)" -ForegroundColor Green
} else {
    Write-Host "  --> [MISSING] .NET Framework 4.8+ not detected. Auto-downloading official installer..." -ForegroundColor Magenta
    $netInstaller = Join-Path $env:TEMP "ndp48-web.exe"
    try {
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?LinkId=2085155" -OutFile $netInstaller -UseBasicParsing
        Write-Host "  --> [INSTALLING] Launching .NET Framework 4.8 setup..." -ForegroundColor Cyan
        Start-Process -FilePath $netInstaller -ArgumentList "/promptrestart" -Wait
        Write-Host "  --> [DONE] .NET Framework setup completed." -ForegroundColor Green
    } catch {
        Write-Host "  --> [ERROR] Download failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Download manually: https://go.microsoft.com/fwlink/?linkid=2088631" -ForegroundColor Yellow
    }
}

Write-Host ""

# 2. Check Edge WebView2 Runtime
Write-Host "[2/3] Checking Microsoft Edge WebView2 Runtime..." -ForegroundColor Yellow
$wv2Installed = $false
$wv2Version = ""

$wv2Folders = @(
    "C:\Program Files (x86)\Microsoft\EdgeWebView\Application",
    "C:\Program Files\Microsoft\EdgeWebView\Application",
    (Join-Path $env:LOCALAPPDATA "Microsoft\EdgeWebView\Application")
)
foreach ($f in $wv2Folders) {
    if (Test-Path $f) {
        $verDirs = Get-ChildItem -Path $f -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "^\d+\.\d+\.\d+\.\d+$" }
        if ($verDirs) {
            $wv2Installed = $true
            $wv2Version = ($verDirs | Select-Object -First 1).Name
            break
        }
    }
}

if (!$wv2Installed) {
    $wv2Keys = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-F500-4414-A21F-42A6A80BE50E}",
        "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F500-4414-A21F-42A6A80BE50E}",
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F500-4414-A21F-42A6A80BE50E}"
    )
    foreach ($k in $wv2Keys) {
        if (Test-Path $k) {
            $pv = (Get-ItemProperty -Path $k -ErrorAction SilentlyContinue).pv
            if ($pv -and $pv -ne "0.0.0.0") {
                $wv2Installed = $true
                $wv2Version = $pv
                break
            }
        }
    }
}

if ($wv2Installed) {
    Write-Host "  --> [OK] Microsoft Edge WebView2 Runtime detected (Version: $wv2Version)" -ForegroundColor Green
} else {
    Write-Host "  --> [MISSING] WebView2 Runtime not detected. Auto-downloading official installer..." -ForegroundColor Magenta
    $wv2Installer = Join-Path $env:TEMP "MicrosoftEdgeWebview2Setup.exe"
    try {
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile $wv2Installer -UseBasicParsing
        Write-Host "  --> [INSTALLING] Installing Microsoft Edge WebView2 Runtime silently..." -ForegroundColor Cyan
        $proc = Start-Process -FilePath $wv2Installer -ArgumentList "/silent /install" -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq -2147219416) {
            Write-Host "  --> [OK] Microsoft Edge WebView2 Runtime ready!" -ForegroundColor Green
            $wv2Installed = $true
        } else {
            Write-Host "  --> [INFO] Installer exited with code: $($proc.ExitCode)." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  --> [ERROR] Download failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Download manually: https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Yellow
    }
}

Write-Host ""

# 3. Check Office / WPS Host
Write-Host "[3/3] Checking Office PowerPoint / WPS Presentation hosts..." -ForegroundColor Yellow
$pptInstalled = $false
$wpsInstalled = $false

$pptKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\PowerPoint\InstallRoot",
    "HKLM:\SOFTWARE\Microsoft\Office\15.0\PowerPoint\InstallRoot",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\PowerPoint\InstallRoot",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\PowerPoint\InstallRoot"
)
foreach ($k in $pptKeys) {
    if (Test-Path $k) { $pptInstalled = $true; break }
}

$wpsKeys = @(
    "HKCU:\SOFTWARE\Kingsoft\Office",
    "HKLM:\SOFTWARE\Kingsoft\Office",
    "HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office"
)
foreach ($k in $wpsKeys) {
    if (Test-Path $k) { $wpsInstalled = $true; break }
}

if ($pptInstalled) {
    Write-Host "  --> [OK] Microsoft Office PowerPoint detected." -ForegroundColor Green
}
if ($wpsInstalled) {
    Write-Host "  --> [OK] WPS Presentation (WPP) detected." -ForegroundColor Green
}
if (!$pptInstalled -and !$wpsInstalled) {
    Write-Host "  --> [INFO] Standard registry key not found (custom path installation is also supported)." -ForegroundColor Cyan
}

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "             Environment check & repair finished!                " -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Cyan
