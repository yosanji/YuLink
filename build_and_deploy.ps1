$ProgressPreference = 'SilentlyContinue'

# Dynamically locate the source directory based on the script root
$srcDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($srcDir)) {
    $srcDir = "."
}

$tempDir = Join-Path $srcDir "PPTBuildTemp"
$extractedDir = Join-Path $srcDir "PowerPointAddInExtracted"
$outDll = "$extractedDir\PowerPointAddIn.dll"

# Ensure PowerPointAddInExtracted directory exists
if (!(Test-Path $extractedDir)) {
    New-Item -ItemType Directory -Path $extractedDir -Force | Out-Null
}

# =========================================================================
# Pre-Flight Dependency Diagnostic & Auto-Fix
# =========================================================================
Write-Host "Checking system dependencies..." -ForegroundColor Cyan

# 1. Check .NET Framework 4.8+
$net48Installed = $false
try {
    $release = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
    if ($release -and $release -ge 528040) { # 528040 = .NET 4.8
        $net48Installed = $true
    }
} catch {}

if ($net48Installed) {
    Write-Host "  [OK] .NET Framework 4.8+ detected." -ForegroundColor Green
} else {
    Write-Host "  [AUTO-FIX] .NET Framework 4.8+ missing. Downloading official installer..." -ForegroundColor Yellow
    $netInstaller = Join-Path $env:TEMP "ndp48-web.exe"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?LinkId=2085155" -OutFile $netInstaller -UseBasicParsing
        Write-Host "  [AUTO-FIX] Launching .NET Framework 4.8 setup..." -ForegroundColor Yellow
        Start-Process -FilePath $netInstaller -ArgumentList "/promptrestart" -Wait
    } catch {
        Write-Host "  [WARN] Failed to auto-download .NET: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "         Please install manually from: https://go.microsoft.com/fwlink/?linkid=2088631" -ForegroundColor Yellow
    }
}

# 2. Check Edge WebView2 Runtime
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
    Write-Host "  [OK] Microsoft Edge WebView2 Runtime detected ($wv2Version)." -ForegroundColor Green
} else {
    Write-Host "  [AUTO-FIX] WebView2 Runtime missing. Downloading official installer..." -ForegroundColor Yellow
    $wv2Installer = Join-Path $env:TEMP "MicrosoftEdgeWebview2Setup.exe"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile $wv2Installer -UseBasicParsing
        Write-Host "  [AUTO-FIX] Installing Microsoft Edge WebView2 Runtime silently..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath $wv2Installer -ArgumentList "/silent /install" -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq -2147219416) {
            Write-Host "  [OK] WebView2 Runtime installed successfully!" -ForegroundColor Green
        }
    } catch {
        Write-Host "  [WARN] Failed to auto-install WebView2: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "         Please install manually from: https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Yellow
    }
}
# =========================================================================


Write-Host "Creating temp directory..."
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir -Force

Write-Host "Copying and rewriting C# source files with correct namespace (UTF8)..."
Get-ChildItem -Path $srcDir -Filter "*.cs" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw -Encoding UTF8
    $content = $content -replace "namespace PPTWebBrowserAddIn", "namespace PowerPointAddIn"
    $content = $content -replace "using PPTWebBrowserAddIn", "using PowerPointAddIn"
    $content = $content -replace "PPTWebBrowserAddIn.Ribbon.xml", "PowerPointAddIn.Ribbon.xml"
    
    $targetPath = Join-Path $tempDir $_.Name
    $content | Set-Content $targetPath -Force -Encoding UTF8
}

$propertiesSrcDir = Join-Path $srcDir "Properties"
if (Test-Path $propertiesSrcDir) {
    $propertiesTempDir = Join-Path $tempDir "Properties"
    New-Item -ItemType Directory -Path $propertiesTempDir -Force
    Get-ChildItem -Path $propertiesSrcDir -Filter "*.cs" | ForEach-Object {
        $content = Get-Content $_.FullName -Raw -Encoding UTF8
        $content = $content -replace "namespace PPTWebBrowserAddIn", "namespace PowerPointAddIn"
        $content = $content -replace "using PPTWebBrowserAddIn", "using PowerPointAddIn"
        $targetPath = Join-Path $propertiesTempDir $_.Name
        $content | Set-Content $targetPath -Force -Encoding UTF8
    }
}

Write-Host "Locating compiler..."
$frameworkDir = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
$csc = Join-Path $frameworkDir "csc.exe"
if (!(Test-Path $csc)) {
    Write-Error "C# Compiler (csc.exe) not found at $csc!"
    exit 1
}

Write-Host "Compiling custom VSTO PowerPointAddIn.dll..."
$refs = @(
    (Join-Path $frameworkDir "System.dll"),
    (Join-Path $frameworkDir "System.Data.dll"),
    (Join-Path $frameworkDir "System.Drawing.dll"),
    (Join-Path $frameworkDir "System.Windows.Forms.dll"),
    (Join-Path $frameworkDir "System.Xml.dll"),
    (Join-Path $frameworkDir "System.Core.dll"),
    (Join-Path $frameworkDir "Microsoft.CSharp.dll"),
    (Join-Path $frameworkDir "System.Net.Http.dll"),
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll",
    "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll",
    "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Office.Tools.Common\v4.0_10.0.0.0__b03f5f7f11d50a3a\Microsoft.Office.Tools.Common.dll",
    "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Office.Tools.v4.0.Framework\v4.0_10.0.0.0__b03f5f7f11d50a3a\Microsoft.Office.Tools.v4.0.Framework.dll",
    "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Office.Tools\v4.0_10.0.0.0__b03f5f7f11d50a3a\Microsoft.Office.Tools.dll",
    "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.VisualStudio.Tools.Applications.Runtime\v4.0_10.0.0.0__b03f5f7f11d50a3a\Microsoft.VisualStudio.Tools.Applications.Runtime.dll",
    (Join-Path $extractedDir "Microsoft.Office.Tools.Common.v4.0.Utilities.dll"),
    (Join-Path $extractedDir "Microsoft.Web.WebView2.Core.dll"),
    (Join-Path $extractedDir "Microsoft.Web.WebView2.WinForms.dll"),
    (Join-Path $extractedDir "QRCoder.dll")
)

$refArgs = $refs | ForEach-Object { "/r:`"$_`"" }
$sourceFiles = @()
Get-ChildItem -Path $tempDir -Filter "*.cs" -Recurse | ForEach-Object {
    $sourceFiles += "`"$($_.FullName)`""
}

$ribbonXml = Join-Path $srcDir "Ribbon.xml"

# Run compilation
$cmdArgs = @(
    "/target:library",
    "/out:`"$outDll`"",
    "/resource:`"$ribbonXml`,PowerPointAddIn.Ribbon.xml`""
) + $refArgs + $sourceFiles

$cmdLine = "$csc " + ($cmdArgs -join " ")
Write-Host "Running: csc..."
Invoke-Expression $cmdLine

if ($LASTEXITCODE -ne 0) {
    Write-Error "Compilation failed!"
    exit 1
}
Write-Host "Compilation successful! DLL generated at $outDll"

# Clean up build temp
Remove-Item $tempDir -Recurse -Force

# =========================================================================
# Self-Signed Certificate Generation & ClickOnce Re-Signing
# =========================================================================
Write-Host "Setting up code-signing certificate..."

$certSubject = "CN=PowerPointAddInLocal"
$pfxPath = "$extractedDir\temp_signing.pfx"
$passwordStr = "Password123"
$password = ConvertTo-SecureString $passwordStr -AsPlainText -Force

# 1. Clean old local certs
Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $certSubject } | ForEach-Object {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $store.Remove($_)
    $store.Close()
}

# 2. Generate new CodeSigning certificate (using Microsoft Enhanced RSA and AES Cryptographic Provider)
$cert = New-SelfSignedCertificate -Subject $certSubject -Type CodeSigning -CertStoreLocation Cert:\CurrentUser\My -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"
Write-Host "Code signing certificate generated successfully."

# 3. Export to PFX
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null
Write-Host "PFX exported to $pfxPath."

# 4. Extract public key XML
$rsa = [System.Security.Cryptography.RSACryptoServiceProvider]$cert.PublicKey.Key
$publicKeyXml = $rsa.ToXmlString($false)
Write-Host "Public key XML extracted successfully."

# 5. Sign the manifests using .NET MSBuild Deployment utilities
Write-Host "Signing ClickOnce manifests..."

function Get-FileHashBase64($path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $hashBytes = $sha.ComputeHash($bytes)
    return [System.Convert]::ToBase64String($hashBytes)
}

$manifestPath = "$extractedDir\PowerPointAddIn.dll.manifest"
$vstoPath = "$extractedDir\PowerPointAddIn.vsto"

# Sign dll.manifest
if (Test-Path $manifestPath) {
    $mXml = [xml](Get-Content $manifestPath -Raw -Encoding UTF8)
    $mAssembly = $mXml.SelectSingleNode("//*[local-name()='assembly']")

    # Remove old signature
    $sig = $mAssembly.SelectSingleNode("*[local-name()='Signature']")
    if ($sig) {
        $mAssembly.RemoveChild($sig) | Out-Null
    }

    # Find the dependentAssembly for PowerPointAddIn.dll
    $dllDep = $mAssembly.SelectSingleNode("//*[local-name()='dependentAssembly' and *[local-name()='assemblyIdentity' and @name='PowerPointAddIn']]")
    if ($dllDep) {
        $dllSize = (Get-Item $outDll).Length
        $dllHash = Get-FileHashBase64 $outDll
        $dllDep.SetAttribute("size", $dllSize.ToString())
        $digest = $dllDep.SelectSingleNode(".//*[local-name()='DigestValue']")
        if ($digest) {
            $digest.InnerText = $dllHash
        }
    }
    $mXml.Save($manifestPath)

    # Perform SignFile via MSBuild Tasks DLL
    Add-Type -Path "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Microsoft.Build.Tasks.v4.0.dll"
    [Microsoft.Build.Tasks.Deployment.ManifestUtilities.SecurityUtilities]::SignFile($pfxPath, $password, $null, $manifestPath)
    Write-Host "Signed PowerPointAddIn.dll.manifest successfully."
}

# Sign vsto
if (Test-Path $vstoPath) {
    # Read the updated publicKeyToken generated by SignFile from dll.manifest
    $mXmlSigned = [xml](Get-Content $manifestPath -Raw -Encoding UTF8)
    $token = $mXmlSigned.assembly.assemblyIdentity.publicKeyToken
    Write-Host "Manifest publicKeyToken updated to: $token"

    $vXml = [xml](Get-Content $vstoPath -Raw -Encoding UTF8)
    $vAssembly = $vXml.SelectSingleNode("//*[local-name()='assembly']")

    # Remove old signature
    $sig2 = $vAssembly.SelectSingleNode("*[local-name()='Signature']")
    if ($sig2) {
        $vAssembly.RemoveChild($sig2) | Out-Null
    }

    # Update dependent manifest identity and hash
    $manifestDep = $vAssembly.SelectSingleNode("//*[local-name()='dependentAssembly' and *[local-name()='assemblyIdentity' and @name='PowerPointAddIn.dll']]")
    if ($manifestDep) {
        $depIdent = $manifestDep.SelectSingleNode("*[local-name()='assemblyIdentity']")
        if ($depIdent) {
            $depIdent.SetAttribute("publicKeyToken", $token) | Out-Null
        }

        $manifestSize = (Get-Item $manifestPath).Length
        $manifestHash = Get-FileHashBase64 $manifestPath
        
        $manifestDep.SetAttribute("size", $manifestSize.ToString())
        $digest2 = $manifestDep.SelectSingleNode(".//*[local-name()='DigestValue']")
        if ($digest2) {
            $digest2.InnerText = $manifestHash
        }
    }
    $vXml.Save($vstoPath)

    # Sign vsto file
    [Microsoft.Build.Tasks.Deployment.ManifestUtilities.SecurityUtilities]::SignFile($pfxPath, $password, $null, $vstoPath)
    Write-Host "Signed PowerPointAddIn.vsto successfully."
}

# Clean up temp PFX file
if (Test-Path $pfxPath) {
    Remove-Item $pfxPath -Force
}

# =========================================================================
# Inject Trust into User-Level ClickOnce Inclusion List
# =========================================================================
Write-Host "Injecting trust into user-level ClickOnce Inclusion List..."

$inclusionRoot = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
if (!(Test-Path $inclusionRoot)) {
    New-Item -Path $inclusionRoot -Force | Out-Null
}

# Format matching URL
$targetUrl = "file:///$extractedDir/PowerPointAddIn.vsto"
$targetUrl = $targetUrl -replace "\\", "/"

$existingGuid = $null
Get-ChildItem -Path $inclusionRoot | ForEach-Object {
    $urlVal = Get-ItemProperty -Path $_.PSPath -Name "Url" -ErrorAction SilentlyContinue
    if ($urlVal -and $urlVal.Url -eq $targetUrl) {
        $existingGuid = $_.PSChildName
    }
}

if ($existingGuid -eq $null) {
    $existingGuid = "{" + [Guid]::NewGuid().ToString() + "}"
}

$guidPath = Join-Path $inclusionRoot $existingGuid
if (!(Test-Path $guidPath)) {
    New-Item -Path $guidPath -Force | Out-Null
}

Set-ItemProperty -Path $guidPath -Name "Url" -Value $targetUrl -Force
Set-ItemProperty -Path $guidPath -Name "PublicKey" -Value $publicKeyXml -Force
Write-Host "Inclusion entry written to registry: $existingGuid"
# =========================================================================

Write-Host "Registering AddIn in Current User Registry..."
$regPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn"
if (!(Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "YuLink" -Force
Set-ItemProperty -Path $regPath -Name "Description" -Value "YuLink - Modern PowerPoint & WPS Web Embedding VSTO AddIn by yosanji" -Force
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord -Force
Set-ItemProperty -Path $regPath -Name "CommandLineSafe" -Value 1 -Type DWord -Force

$manifestUrl = "file:///$extractedDir/PowerPointAddIn.vsto|vstolocal"
$manifestUrl = $manifestUrl -replace "\\", "/"
Set-ItemProperty -Path $regPath -Name "Manifest" -Value $manifestUrl -Force

Write-Host "Registering AddIn for WPS WPP..."
$wpsPath = "HKCU:\Software\Kingsoft\Office\WPP\AddinsWL"
if (!(Test-Path $wpsPath)) {
    New-Item -Path $wpsPath -Force | Out-Null
}
Set-ItemProperty -Path $wpsPath -Name "PowerPointAddIn" -Value "" -Force

Write-Host "=============================================="
Write-Host "DEPLOYMENT SUCCESSFUL! You can now launch PowerPoint or WPS."
Write-Host "=============================================="
