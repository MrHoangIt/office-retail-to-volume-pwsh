<#
.SYNOPSIS
    Office 2019 Retail to Volume License Converter
.DESCRIPTION
    Converts Office 2019 Retail installation to Volume License with error handling, detection, and KMS-only license file installation
.NOTES
    Version: 0.3
	Create date: 27-August-2025
	Last update: 27-August-2025
    Requires: PowerShell 5.1+ and Administrator privileges
	Run "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" if current policy is "restricted"
.AUTHOR
	Harry Hoang Le
	Phone/Zalo/Whatsapp: +84 888441779
	(this is a Vibe-coding script)
#>

# Set encoding and error handling
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$ErrorActionPreference = 'Continue'

# Global variables
$script:LogMessages = @()
$script:DebugMode = $false

#region Utility Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    $script:LogMessages += $logEntry
    
    switch ($Level) {
        'Info' { Write-Host $Message -ForegroundColor White }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error' { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
    }
    
    if ($script:DebugMode) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
    }
}

function Save-Log {
    $logPath = Join-Path $env:TEMP "OfficeConverter_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $script:LogMessages | Out-File -FilePath $logPath -Encoding UTF8
    Write-Log "Log saved to: $logPath" -Level Info
}

function Test-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PowerShellVersion {
    $version = $PSVersionTable.PSVersion
    if ($version.Major -lt 5) {
        Write-Log "PowerShell version $($version.ToString()) is not supported. Requires 5.1 or higher." -Level Error
        return $false
    }
    return $true
}

function Set-ExecutionPolicyIfNeeded {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser -ErrorAction SilentlyContinue
    
    if ($currentPolicy -notin @('RemoteSigned', 'Unrestricted', 'Bypass')) {
        try {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            Write-Log "Execution Policy set to RemoteSigned for CurrentUser." -Level Success
            return $true
        }
        catch {
            Write-Log "Failed to set Execution Policy: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    
    Write-Log "Execution Policy is already configured appropriately: $currentPolicy" -Level Info
    return $true
}
#endregion

#region System Validation
function Test-WindowsEdition {
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $osCaption = $osInfo.Caption
        
        Write-Log "Detected Windows: $osCaption" -Level Info
        
        if ($osCaption -match "Home") {
            Write-Log "Windows Home edition does not support Volume Licensing. Please upgrade to Pro or Enterprise." -Level Error
            return $false
        }
        
        $version = [System.Environment]::OSVersion.Version
        if ($version.Major -lt 10) {
            Write-Log "Windows version $($version.ToString()) may have compatibility issues." -Level Warning
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to detect Windows edition: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Get-OfficeArchitecture {
    $architectures = @()
    
    $regPaths = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun"; Arch = "64-bit" },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun"; Arch = "32-bit" }
    )
    
    foreach ($regPath in $regPaths) {
        if (Test-Path $regPath.Path) {
            try {
                $config = Get-ItemProperty -Path "$($regPath.Path)\Configuration" -ErrorAction Stop
                $platform = $config.Platform
                $version = $config.VersionToReport
                
                Write-Log "Found Office registration: $($regPath.Arch), Platform: $platform, Version: $version" -Level Info
                $architectures += $regPath.Arch
            }
            catch {
                Write-Log "Could not read configuration from $($regPath.Path)" -Level Warning
            }
        }
    }
    
    if ($architectures.Count -eq 0) {
        Write-Log "No Office installation found in registry." -Level Error
        return $null
    }
    
    return $architectures[0]
}

function Get-OfficeInstallPath {
    param([string]$Architecture)
    
    $possiblePaths = @()
    
    if ($Architecture -eq "64-bit") {
        $possiblePaths += "$env:ProgramFiles\Microsoft Office\Office16"
    } else {
        $possiblePaths += "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
    }
    
    try {
        $regPath = if ($Architecture -eq "32-bit") { 
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
        } else {
            "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        }
        
        if (Test-Path $regPath) {
            $config = Get-ItemProperty -Path $regPath -ErrorAction Stop
            if ($config.InstallationPath) {
                $possiblePaths += "$($config.InstallationPath)\Office16"
            }
        }
    }
    catch {
        Write-Log "Could not read installation path from registry: $($_.Exception.Message)" -Level Warning
    }
    
    foreach ($path in $possiblePaths) {
        if ($path -and (Test-Path "$path\ospp.vbs")) {
            Write-Log "Found Office installation at: $path" -Level Success
            return $path
        }
    }
    
    Write-Log "Could not locate Office installation. Checked paths: $($possiblePaths -join ', ')" -Level Error
    return $null
}
#endregion

#region License File Installation
function Install-LicenseFiles {
    param(
        [string]$ScriptPath,
        [string]$ProductType,
        [string]$OfficePath
    )
    
    $licenseFolder = Join-Path (Split-Path $OfficePath -Parent) "root\Licenses16"
    $licensePattern = switch ($ProductType) {
        "ProPlus2019" { "ProPlus2019VL_KMS_Client*.xrm-ms" }
        "VisioPro2019" { "VisioPro2019VL_KMS_Client*.xrm-ms" }
        "ProjectPro2019" { "ProjectPro2019VL_KMS_Client*.xrm-ms" }
        default { $null }
    }
    
    if (-not $licensePattern) {
        Write-Log "No license file pattern defined for product type: $ProductType" -Level Error
        return $false
    }
    
    if (-not (Test-Path $licenseFolder)) {
        Write-Log "License folder not found at: $licenseFolder" -Level Error
        Write-Log "Please download Office Deployment Tool (ODT) from https://www.microsoft.com/en-us/download/details.aspx?id=49117" -Level Warning
        Write-Log "Use ODT to download Office 2019 Volume License package and extract the 'root\Licenses16' folder to '$licenseFolder'" -Level Warning
        Write-Log "Example configuration.xml for ODT:" -Level Info
        Write-Log "<Configuration><Add OfficeClientEdition='64' Channel='PerpetualVL2019'><Product ID='$($ProductType)Volume'><Language ID='en-us' /></Product></Add></Configuration>" -Level Info
        Write-Log "Run: setup.exe /download configuration.xml, then copy 'root\Licenses16' to the specified path." -Level Info
        return $false
    }
    
    Write-Log "Found license folder at: $licenseFolder" -Level Info
    $licenseFiles = Get-ChildItem -Path $licenseFolder -Filter $licensePattern -ErrorAction SilentlyContinue
    
    if ($licenseFiles.Count -eq 0) {
        Write-Log "No KMS license files found for pattern: $licensePattern in $licenseFolder" -Level Error
        return $false
    }
    
    $success = $true
    foreach ($file in $licenseFiles) {
        $licensePath = $file.FullName
        Write-Log "Installing KMS license file: $licensePath" -Level Info
        $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/inslic:`"$licensePath`""
        
        if ($result.Success) {
            Write-Log "Successfully installed KMS license file: $licensePath" -Level Success
        } else {
            Write-Log "Failed to install KMS license file: $licensePath" -Level Error
            $success = $false
        }
    }
    
    return $success
}
#endregion

#region Office Product Detection
function Invoke-OSPPCommand {
    param(
        [string]$ScriptPath,
        [string]$Arguments,
        [switch]$ReturnOutput
    )
    
    try {
        Write-Log "Executing: cscript.exe `"$ScriptPath`" $Arguments" -Level Info
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "cscript.exe"
        $processInfo.Arguments = "`"$ScriptPath`" $Arguments"
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null
        
        $output = $process.StandardOutput.ReadToEnd()
        $error = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        
        if ($process.ExitCode -ne 0) {
            Write-Log "OSPP command failed with exit code: $($process.ExitCode)" -Level Error
            Write-Log "OSPP command error: $error" -Level Error
            return @{ Success = $false; Output = $output; Error = $error }
        }
        
        Write-Log "OSPP command executed successfully" -Level Success
        Write-Log "OSPP command output: $output" -Level Info
        return @{ Success = $true; Output = $output; Error = $error }
    }
    catch {
        Write-Log "Exception executing OSPP command: $($_.Exception.Message)" -Level Error
        return @{ Success = $false; Output = $null; Error = $_.Exception.Message }
    }
}

function Get-InstalledProducts {
    param([string]$ScriptPath)
    
    Write-Log "Retrieving installed Office products information..." -Level Info
    
    $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/dstatusall" -ReturnOutput
    
    if (-not $result.Success) {
        Write-Log "Failed to get product status from ospp.vbs" -Level Error
        return $null
    }
    
    return $result.Output -split "`n"
}

function Get-ProductsFromRegistry {
    param([string]$OfficeArchitecture)
    
    Write-Log "Detecting Office products from registry..." -Level Info
    
    $products = @()
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    
    if ($OfficeArchitecture -eq "32-bit") {
        $regPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    }
    
    if (-not (Test-Path $regPath)) {
        Write-Log "Registry path not found: $regPath" -Level Warning
        return $products
    }
    
    try {
        $config = Get-ItemProperty -Path $regPath -ErrorAction Stop
        $productReleaseIds = $config.ProductReleaseIds
        
        if (-not $productReleaseIds) {
            Write-Log "No ProductReleaseIds found in registry" -Level Warning
            return $products
        }
        
        Write-Log "Found ProductReleaseIds: $productReleaseIds" -Level Info
        
        $productMappings = @{
            "ProPlus2019Retail" = @{ Type = "ProPlus2019"; License = "Retail"; NeedsConversion = $true }
            "ProPlus2019Volume" = @{ Type = "ProPlus2019"; License = "Volume"; NeedsConversion = $false }
            "VisioPro2019Retail" = @{ Type = "VisioPro2019"; License = "Retail"; NeedsConversion = $true }
            "VisioPro2019Volume" = @{ Type = "VisioPro2019"; License = "Volume"; NeedsConversion = $false }
            "ProjectPro2019Retail" = @{ Type = "ProjectPro2019"; License = "Retail"; NeedsConversion = $true }
            "ProjectPro2019Volume" = @{ Type = "ProjectPro2019"; License = "Volume"; NeedsConversion = $false }
        }
        
        foreach ($mapping in $productMappings.GetEnumerator()) {
            if ($productReleaseIds -match $mapping.Key) {
                $product = @{
                    Type = $mapping.Value.Type
                    License = $mapping.Value.License
                    NeedsConversion = $mapping.Value.NeedsConversion
                    PartialKey = $null
                }
                $products += $product
                $status = if ($mapping.Value.NeedsConversion) { "NEEDS CONVERSION" } else { "OK" }
                Write-Log "Detected $($mapping.Key): $status" -Level Info
            }
        }
        
        return $products
    }
    catch {
        Write-Log "Error reading registry configuration: $($_.Exception.Message)" -Level Error
        return $products
    }
}

function Parse-OSPPOutput {
    param([array]$Output)
    
    if (-not $Output) {
        Write-Log "No output to parse" -Level Warning
        return @()
    }
    
    $products = @()
    $currentProduct = $null
    $currentLicense = $null
    $currentDescription = $null
    
    foreach ($line in $Output) {
        $line = $line.Trim()
        
        if ($line -match "LICENSE NAME:\s*(.*)") {
            $currentLicense = $matches[1].Trim()
            Write-Log "Found license name: $currentLicense" -Level Info
            
            if ($currentLicense -match "Office19ProPlus2019") {
                $currentProduct = "ProPlus2019"
            }
            elseif ($currentLicense -match "VisioPro2019") {
                $currentProduct = "VisioPro2019"
            }
            elseif ($currentLicense -match "ProjectPro2019") {
                $currentProduct = "ProjectPro2019"
            }
        }
        elseif ($line -match "LICENSE DESCRIPTION:\s*(.*)") {
            $currentDescription = $matches[1].Trim()
            Write-Log "Found license description: $currentDescription" -Level Info
            
            if (-not $currentProduct) {
                if ($currentDescription -match "Office.*Professional Plus.*2019|Office.*19.*RETAIL") {
                    $currentProduct = "ProPlus2019"
                }
                elseif ($currentDescription -match "Visio.*Professional.*2019") {
                    $currentProduct = "VisioPro2019"
                }
                elseif ($currentDescription -match "Project.*Professional.*2019") {
                    $currentProduct = "ProjectPro2019"
                }
            }
        }
        elseif ($line -match "Last 5 characters of installed product key:\s*(.*)") {
            if ($currentProduct) {
                $partialKey = $matches[1].Trim()
                
                $needsConversion = $true
                if ($currentLicense -match "VOLUME" -or $currentDescription -match "VOLUME") {
                    $needsConversion = $false
                }
                elseif ($currentDescription -match "RETAIL|Grace") {
                    $needsConversion = $true
                }
                
                $licenseType = if ($needsConversion) { "Retail" } else { "Volume" }
                
                $product = @{
                    Type = $currentProduct
                    License = $licenseType
                    NeedsConversion = $needsConversion
                    PartialKey = $partialKey
                }
                
                $products += $product
                Write-Log "Found product: $currentProduct ($licenseType) - Key: $partialKey - Needs conversion: $needsConversion" -Level Info
                
                $currentProduct = $null
                $currentLicense = $null
                $currentDescription = $null
            }
            else {
                Write-Log "Found partial key $($matches[1].Trim()) but no product identified" -Level Warning
            }
        }
    }
    
    return $products
}

function Get-OfficeProductsAdvanced {
    param(
        [string]$ScriptPath,
        [string]$OfficeArchitecture
    )
    
    $allProducts = @()
    
    Write-Log "Retrieving all Office products for detailed status..." -Level Info
    $osppOutput = Get-InstalledProducts -ScriptPath $ScriptPath
    if ($osppOutput) {
        $osppProducts = Parse-OSPPOutput -Output $osppOutput
        if ($osppProducts.Count -gt 0) {
            Write-Log "Found $($osppProducts.Count) products via OSPP" -Level Success
            foreach ($product in $osppProducts) {
                Write-Log "Product: $($product.Type), License: $($product.License), Key: $($product.PartialKey), NeedsConversion: $($product.NeedsConversion)" -Level Info
            }
            $allProducts += $osppProducts
        }
    }
    
    if ($allProducts.Count -eq 0) {
        Write-Log "No products found via OSPP, checking registry..." -Level Warning
        $regProducts = Get-ProductsFromRegistry -OfficeArchitecture $OfficeArchitecture
        if ($regProducts.Count -gt 0) {
            Write-Log "Found $($regProducts.Count) products via registry" -Level Success
            foreach ($product in $regProducts) {
                Write-Log "Product: $($product.Type), License: $($product.License), NeedsConversion: $($product.NeedsConversion)" -Level Info
            }
            $allProducts += $regProducts
        }
    }
    
    $uniqueProducts = $allProducts | Sort-Object -Property Type -Unique
    return $uniqueProducts
}
#endregion

#region License Management
function Remove-RetailLicense {
    param(
        [string]$ScriptPath,
        [string]$PartialKey
    )
    
    if (-not $PartialKey) {
        Write-Log "No partial key provided for removal" -Level Warning
        return $true
    }
    
    Write-Log "Removing retail license with key ending in: $PartialKey" -Level Info
    
    $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/unpkey:$PartialKey"
    
    if ($result.Success) {
        Write-Log "Retail license removed successfully" -Level Success
        
        # Reset license state
        Write-Log "Resetting license state with /rearm..." -Level Info
        $rearmResult = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/rearm"
        if ($rearmResult.Success) {
            Write-Log "License state reset successfully" -Level Success
        } else {
            Write-Log "Failed to reset license state: $($rearmResult.Error)" -Level Warning
        }
        return $true
    } else {
        Write-Log "Failed to remove retail license" -Level Error
        return $false
    }
}

function Install-VolumeLicense {
    param(
        [string]$ScriptPath,
        [string]$ProductType,
        [string]$OfficePath
    )
    
    $kmsKeys = @{
        "ProPlus2019" = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
        "VisioPro2019" = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
        "ProjectPro2019" = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
    }
    
    $key = $kmsKeys[$ProductType]
    if (-not $key) {
        Write-Log "No KMS key available for product type: $ProductType" -Level Error
        return $false
    }
    
    # Install KMS license files
    Write-Log "Checking and installing KMS Volume License files for $ProductType" -Level Info
    if (-not (Install-LicenseFiles -ScriptPath $ScriptPath -ProductType $ProductType -OfficePath $OfficePath)) {
        Write-Log "Failed to install KMS Volume License files for $ProductType" -Level Error
        return $false
    }
    
    Write-Log "Installing volume license for $ProductType with key: $key" -Level Info
    
    $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/inpkey:$key"
    
    if ($result.Success) {
        Write-Log "Volume license installed successfully for $ProductType" -Level Success
        
        # Verify key installation
        Write-Log "Verifying key installation for $ProductType..." -Level Info
        $verifyResult = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/dstatus" -ReturnOutput
        if ($verifyResult.Success) {
            if ($verifyResult.Output -match $key.Substring($key.Length - 5)) {
                Write-Log "Key verification successful: KMS key ending in $($key.Substring($key.Length - 5)) detected" -Level Success
            } else {
                Write-Log "Key verification failed: KMS key not detected in /dstatus output" -Level Error
                return $false
            }
        } else {
            Write-Log "Failed to verify key installation: $($verifyResult.Error)" -Level Error
            return $false
        }
        return $true
    } else {
        Write-Log "Failed to install volume license for $ProductType" -Level Error
        return $false
    }
}

function Set-KMSServer {
    param([string]$ScriptPath)
    
    Write-Host "`nKMS Server Configuration:" -ForegroundColor Yellow
    Write-Host "1. Enter custom KMS server"
    Write-Host "2. Skip KMS server configuration"
    
    do {
        $choice = Read-Host "`nEnter your choice (1-2)"
        switch ($choice) {
            "1" {
                $kmsServer = Read-Host "Enter KMS server address"
                if ($kmsServer) {
                    Write-Log "Setting KMS server to: $kmsServer" -Level Info
                    
                    $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/sethst:$kmsServer"
                    if ($result.Success) {
                        Write-Log "KMS server configured successfully" -Level Success
                        return $true
                    } else {
                        Write-Log "Failed to configure KMS server" -Level Error
                        return $false
                    }
                }
                return $true
            }
            "2" {
                Write-Log "Skipping KMS server configuration" -Level Info
                Write-Log "Warning: Skipping KMS server may prevent successful activation" -Level Warning
                return $true
            }
            default {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        }
    } while ($true)
}

function Start-OfficeActivation {
    param([string]$ScriptPath)
    
    Write-Host "`nOffice Activation:" -ForegroundColor Yellow
    Write-Host "1. Activate Office now"
    Write-Host "2. Skip activation"
    
    do {
        $choice = Read-Host "`nEnter your choice (1-2)"
        switch ($choice) {
            "1" {
                Write-Log "Attempting to activate Office..." -Level Info
                
                $result = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/act"
                if ($result.Success) {
                    Write-Log "Office activation attempted successfully" -Level Success
                    # Verify activation status
                    $statusResult = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/dstatusall" -ReturnOutput
                    if ($statusResult.Success) {
                        if ($statusResult.Output -match "LICENSE STATUS:.*LICENSED") {
                            Write-Log "Final verification: Office is fully activated (LICENSED)" -Level Success
                        } else {
                            Write-Log "Final verification: Office activation completed but not fully LICENSED. Check /dstatusall output for details." -Level Warning
                            Write-Log "Current status output: $($statusResult.Output)" -Level Info
                        }
                    } else {
                        Write-Log "Failed to verify activation status: $($statusResult.Error)" -Level Error
                    }
                    return $true
                } else {
                    Write-Log "Office activation failed: $($result.Error)" -Level Error
                    return $false
                }
            }
            "2" {
                Write-Log "Skipping Office activation" -Level Info
                Write-Log "Warning: Skipping activation may leave Office in an unlicensed state" -Level Warning
                return $true
            }
            default {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        }
    } while ($true)
}
#endregion

#region Main Script
function Show-Banner {
    Write-Host @"
╔══════════════════════════════════════════════════════════════════╗
║           Office 2019 Retail to Volume License Converter         ║
║                      Version 0.3 (27-Aug-2025)                   ║
║                       Author: Harry Hoang Le                     ║
╚══════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
}

function Test-Prerequisites {
    Write-Log "Checking system prerequisites..." -Level Info
    
    if (-not (Test-AdminRights)) {
        Write-Log "Administrator privileges required. Please run PowerShell as Administrator." -Level Error
        return $false
    }
    Write-Log "Administrator privileges confirmed" -Level Success
    
    if (-not (Test-PowerShellVersion)) {
        return $false
    }
    Write-Log "PowerShell version check passed" -Level Success
    
    if (-not (Set-ExecutionPolicyIfNeeded)) {
        return $false
    }
    
    if (-not (Test-WindowsEdition)) {
        return $false
    }
    
    Write-Log "All prerequisites met" -Level Success
    return $true
}

function Start-ConversionProcess {
    Write-Log "Starting Office 2019 conversion process..." -Level Info
    
    $officeArch = Get-OfficeArchitecture
    if (-not $officeArch) {
        Write-Log "Could not detect Office installation" -Level Error
        return $false
    }
    Write-Log "Detected Office architecture: $officeArch" -Level Success
    
    $officePath = Get-OfficeInstallPath -Architecture $officeArch
    if (-not $officePath) {
        Write-Log "Could not locate Office installation path" -Level Error
        return $false
    }
    
    $scriptPath = Join-Path $officePath "ospp.vbs"
    Write-Log "Using OSPP script: $scriptPath" -Level Info
    
    $products = Get-OfficeProductsAdvanced -ScriptPath $scriptPath -OfficeArchitecture $officeArch
    
    if ($products.Count -eq 0) {
        Write-Log "No Office 2019 products found for conversion" -Level Error
        return $false
    }
    
    Write-Log "Found $($products.Count) Office products" -Level Info
    
    foreach ($product in $products) {
        $status = if ($product.NeedsConversion) { "NEEDS CONVERSION" } else { "OK (Volume)" }
        Write-Log "Product: $($product.Type) - License: $($product.License) - Key: $($product.PartialKey) - Status: $status" -Level Info
    }
    
    $productsToConvert = $products | Where-Object { $_.NeedsConversion -eq $true }
    
    if ($productsToConvert.Count -eq 0) {
        Write-Log "All detected products are already using Volume licensing" -Level Success
        return $true
    }
    
    Write-Log "Products requiring conversion: $($productsToConvert.Count)" -Level Info
    
    Write-Log "Getting current license status for key removal..." -Level Info
    $currentStatus = Get-InstalledProducts -ScriptPath $scriptPath
    $currentProducts = Parse-OSPPOutput -Output $currentStatus
    
    $conversionSuccess = $true
    foreach ($product in $productsToConvert) {
        Write-Log "Processing product: $($product.Type)" -Level Info
        
        $retailKey = $null
        if ($product.PartialKey) {
            $retailKey = $product.PartialKey
        }
        else {
            $currentProduct = $currentProducts | Where-Object { $_.Type -eq $product.Type -and $_.NeedsConversion -eq $true }
            if ($currentProduct -and $currentProduct.PartialKey) {
                $retailKey = $currentProduct.PartialKey
            }
        }
        
        if ($retailKey) {
            Write-Log "Found retail key for removal: $retailKey" -Level Info
            if (-not (Remove-RetailLicense -ScriptPath $scriptPath -PartialKey $retailKey)) {
                Write-Log "Failed to remove retail key, but continuing with volume key installation..." -Level Warning
            }
        }
        else {
            Write-Log "No retail key found for removal, installing volume key directly..." -Level Info
        }
        
        if (-not (Install-VolumeLicense -ScriptPath $scriptPath -ProductType $product.Type -OfficePath $officePath)) {
            $conversionSuccess = $false
        }
    }
    
    if (-not $conversionSuccess) {
        Write-Log "Some conversions failed. Check the log for details." -Level Error
        return $false
    }
    
    Write-Log "All products processed. Verifying final status..." -Level Info
    
    $finalStatus = Get-InstalledProducts -ScriptPath $scriptPath
    $finalProducts = Parse-OSPPOutput -Output $finalStatus
    
    foreach ($product in $finalProducts) {
        $status = if ($product.NeedsConversion) { "NEEDS CONVERSION" } else { "OK (Volume)" }
        Write-Log "Final status - Product: $($product.Type) - License: $($product.License) - Key: $($product.PartialKey) - Status: $status" -Level Info
    }
    
    $remainingRetail = $finalProducts | Where-Object { $_.NeedsConversion -eq $true }
    if ($remainingRetail.Count -gt 0) {
        Write-Log "Warning: $($remainingRetail.Count) products still show as Retail. Manual cleanup may be required." -Level Warning
        foreach ($retail in $remainingRetail) {
            Write-Log "Retail product remaining: $($retail.Type) with key: $($retail.PartialKey)" -Level Warning
        }
    }
    else {
        Write-Log "All products successfully converted to Volume licensing" -Level Success
    }
    
    Set-KMSServer -ScriptPath $scriptPath
    Start-OfficeActivation -ScriptPath $scriptPath
    
    return $true
}

function Main {
    Show-Banner
    
    try {
        if (-not (Test-Prerequisites)) {
            Write-Log "Prerequisites not met. Exiting." -Level Error
            return
        }
        
        $success = Start-ConversionProcess
        
        if ($success) {
            Write-Log "Conversion process completed successfully!" -Level Success
        } else {
            Write-Log "Conversion process completed with errors. Check the log for details." -Level Error
        }
    }
    catch {
        Write-Log "Unexpected error occurred: $($_.Exception.Message)" -Level Error
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    }
    finally {
        Save-Log
        Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
        Read-Host
    }
}

if ($args -contains "-Debug") {
    $script:DebugMode = $true
}

Main
#endregion