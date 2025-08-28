<#
.SYNOPSIS
    Office 2019/2021/2024 Retail to Volume License Converter
.DESCRIPTION
    Converts Office 2019, 2021, and 2024 Retail installation to Volume License with error handling, detection, and KMS-only license file installation
.NOTES
    Version: 0.7
    Create date: 26-August-2025
    Last update: 28-August-2025
    Requires: PowerShell 5.1+ and Administrator privileges
    Run "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force" if current policy is "restricted"
    Author: Harry Hoang Le
    Contact: Phone/Zalo/Whatsapp: +84 888441779
    (This is a Vibe-coding script)
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

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelaySeconds = 2,
        [string]$OperationName = "Operation"
    )
    
    $attempt = 1
    while ($attempt -le $MaxRetries) {
        try {
            Write-Log "Attempting $OperationName (attempt $attempt/$MaxRetries)" -Level Info
            $result = & $ScriptBlock
            Write-Log "$OperationName succeeded on attempt $attempt" -Level Success
            return $result
        }
        catch {
            Write-Log "$OperationName failed on attempt $attempt : $($_.Exception.Message)" -Level Warning
            if ($attempt -eq $MaxRetries) {
                Write-Log "$OperationName failed after $MaxRetries attempts" -Level Error
                throw
            }
            Start-Sleep -Seconds $DelaySeconds
            $attempt++
        }
    }
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
    
    # Check for different Office versions
    $officeVersions = @("Office16", "Office15", "Office14")
    
    foreach ($version in $officeVersions) {
        if ($Architecture -eq "64-bit") {
            $possiblePaths += "$env:ProgramFiles\Microsoft Office\$version"
        } else {
            $possiblePaths += "${env:ProgramFiles(x86)}\Microsoft Office\$version"
        }
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
                foreach ($version in $officeVersions) {
                    $possiblePaths += "$($config.InstallationPath)\$version"
                }
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

function Get-OfficeVersion {
    param([string]$OfficePath)
    
    if ($OfficePath -match "Office16") {
        # Could be 2019, 2021, or 2024 - need to check registry for more specific version
        try {
            $config = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
            if (-not $config) {
                $config = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
            }
            
            if ($config.VersionToReport) {
                $version = $config.VersionToReport
                Write-Log "Detected Office version from registry: $version" -Level Info
                
                # Parse version to determine Office edition
                if ($version -match "^16\.0\.1\d{4}\.") {
                    $buildNumber = [int]($version -replace "^16\.0\.(\d{5})\..*", '$1')
                    if ($buildNumber -ge 17928) {
                        return "2024"
                    }
                    elseif ($buildNumber -ge 14332) {
                        return "2021"  
                    }
                    else {
                        return "2019"
                    }
                }
            }
        }
        catch {
            Write-Log "Could not determine specific Office version, defaulting to 2019" -Level Warning
        }
        return "2019"
    }
    elseif ($OfficePath -match "Office15") {
        return "2013"
    }
    elseif ($OfficePath -match "Office14") {
        return "2010"
    }
    
    return "Unknown"
}
#endregion

#region License File Installation
function Get-LicenseFolderPath {
    param(
        [string]$OfficePath,
        [string]$OfficeVersion
    )
    
    $baseFolder = Split-Path $OfficePath -Parent
    
    switch ($OfficeVersion) {
        "2024" { return Join-Path $baseFolder "root\Licenses16" }
        "2021" { return Join-Path $baseFolder "root\Licenses16" }
        "2019" { return Join-Path $baseFolder "root\Licenses16" }
        "2016" { return Join-Path $baseFolder "root\Licenses16" }
        default { return Join-Path $baseFolder "root\Licenses16" }
    }
}

function Get-LicensePattern {
    param(
        [string]$ProductType,
        [string]$OfficeVersion
    )
    
    $patterns = @{
        "2024" = @{
            "ProPlus2024" = "ProPlus2024VL_KMS_Client*.xrm-ms"
            "VisioPro2024" = "VisioPro2024VL_KMS_Client*.xrm-ms"
            "ProjectPro2024" = "ProjectPro2024VL_KMS_Client*.xrm-ms"
        }
        "2021" = @{
            "ProPlus2021" = "ProPlus2021VL_KMS_Client*.xrm-ms"
            "VisioPro2021" = "VisioPro2021VL_KMS_Client*.xrm-ms"
            "ProjectPro2021" = "ProjectPro2021VL_KMS_Client*.xrm-ms"
        }
        "2019" = @{
            "ProPlus2019" = "ProPlus2019VL_KMS_Client*.xrm-ms"
            "VisioPro2019" = "VisioPro2019VL_KMS_Client*.xrm-ms"
            "ProjectPro2019" = "ProjectPro2019VL_KMS_Client*.xrm-ms"
        }
    }
    
    return $patterns[$OfficeVersion][$ProductType]
}

function Install-LicenseFiles {
    param(
        [string]$ScriptPath,
        [string]$ProductType,
        [string]$OfficePath,
        [string]$OfficeVersion
    )
    
    $licenseFolder = Get-LicenseFolderPath -OfficePath $OfficePath -OfficeVersion $OfficeVersion
    $licensePattern = Get-LicensePattern -ProductType $ProductType -OfficeVersion $OfficeVersion
    
    if (-not $licensePattern) {
        Write-Log "No license file pattern defined for product type: $ProductType (Office $OfficeVersion)" -Level Error
        return $false
    }
    
    if (-not (Test-Path $licenseFolder)) {
        Write-Log "License folder not found at: $licenseFolder" -Level Error
        Write-Log "Please download Office Deployment Tool (ODT) from https://www.microsoft.com/en-us/download/details.aspx?id=49117" -Level Warning
        
        $channel = switch ($OfficeVersion) {
            "2024" { "PerpetualVL2024" }
            "2021" { "PerpetualVL2021" }
            "2019" { "PerpetualVL2019" }
            default { "PerpetualVL2019" }
        }
        
        Write-Log "Use ODT to download Office $OfficeVersion Volume License package and extract the 'root\Licenses16' folder to '$licenseFolder'" -Level Warning
        Write-Log "Example configuration.xml for ODT:" -Level Info
        Write-Log "<Configuration><Add OfficeClientEdition='64' Channel='$channel'><Product ID='$($ProductType)Volume'><Language ID='en-us' /></Product></Add></Configuration>" -Level Info
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
        
        $result = Invoke-WithRetry -ScriptBlock {
            Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/inslic:`"$licensePath`""
        } -MaxRetries 3 -OperationName "Install license file $($file.Name)"
        
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
            if ($error) { Write-Log "OSPP command error: $error" -Level Error }
            return @{ Success = $false; Output = $output; Error = $error }
        }
        
        Write-Log "OSPP command executed successfully" -Level Success
        if ($script:DebugMode -and $output) { 
            Write-Log "OSPP command output: $output" -Level Info 
        }
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
    param(
        [string]$OfficeArchitecture,
        [string]$OfficeVersion
    )
    
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
        
        # Enhanced product mappings for multiple Office versions
        $productMappings = @{
            # Office 2024
            "ProPlus2024Retail" = @{ Type = "ProPlus2024"; License = "Retail"; NeedsConversion = $true; Version = "2024" }
            "ProPlus2024Volume" = @{ Type = "ProPlus2024"; License = "Volume"; NeedsConversion = $false; Version = "2024" }
            "VisioPro2024Retail" = @{ Type = "VisioPro2024"; License = "Retail"; NeedsConversion = $true; Version = "2024" }
            "VisioPro2024Volume" = @{ Type = "VisioPro2024"; License = "Volume"; NeedsConversion = $false; Version = "2024" }
            "ProjectPro2024Retail" = @{ Type = "ProjectPro2024"; License = "Retail"; NeedsConversion = $true; Version = "2024" }
            "ProjectPro2024Volume" = @{ Type = "ProjectPro2024"; License = "Volume"; NeedsConversion = $false; Version = "2024" }
            
            # Office 2021
            "ProPlus2021Retail" = @{ Type = "ProPlus2021"; License = "Retail"; NeedsConversion = $true; Version = "2021" }
            "ProPlus2021Volume" = @{ Type = "ProPlus2021"; License = "Volume"; NeedsConversion = $false; Version = "2021" }
            "VisioPro2021Retail" = @{ Type = "VisioPro2021"; License = "Retail"; NeedsConversion = $true; Version = "2021" }
            "VisioPro2021Volume" = @{ Type = "VisioPro2021"; License = "Volume"; NeedsConversion = $false; Version = "2021" }
            "ProjectPro2021Retail" = @{ Type = "ProjectPro2021"; License = "Retail"; NeedsConversion = $true; Version = "2021" }
            "ProjectPro2021Volume" = @{ Type = "ProjectPro2021"; License = "Volume"; NeedsConversion = $false; Version = "2021" }
            
            # Office 2019
            "ProPlus2019Retail" = @{ Type = "ProPlus2019"; License = "Retail"; NeedsConversion = $true; Version = "2019" }
            "ProPlus2019Volume" = @{ Type = "ProPlus2019"; License = "Volume"; NeedsConversion = $false; Version = "2019" }
            "VisioPro2019Retail" = @{ Type = "VisioPro2019"; License = "Retail"; NeedsConversion = $true; Version = "2019" }
            "VisioPro2019Volume" = @{ Type = "VisioPro2019"; License = "Volume"; NeedsConversion = $false; Version = "2019" }
            "ProjectPro2019Retail" = @{ Type = "ProjectPro2019"; License = "Retail"; NeedsConversion = $true; Version = "2019" }
            "ProjectPro2019Volume" = @{ Type = "ProjectPro2019"; License = "Volume"; NeedsConversion = $false; Version = "2019" }
        }
        
        foreach ($mapping in $productMappings.GetEnumerator()) {
            if ($productReleaseIds -match $mapping.Key) {
                $product = @{
                    Type = $mapping.Value.Type
                    License = $mapping.Value.License
                    NeedsConversion = $mapping.Value.NeedsConversion
                    Version = $mapping.Value.Version
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
    param(
        [array]$Output,
        [string]$OfficeVersion
    )
    
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
            
            # Enhanced pattern matching for different Office versions
            if ($currentLicense -match "Office(19|21|24)ProPlus(2019|2021|2024)") {
                $version = $matches[2]
                $currentProduct = "ProPlus$version"
            }
            elseif ($currentLicense -match "VisioPro(2019|2021|2024)") {
                $version = $matches[1]
                $currentProduct = "VisioPro$version"
            }
            elseif ($currentLicense -match "ProjectPro(2019|2021|2024)") {
                $version = $matches[1]
                $currentProduct = "ProjectPro$version"
            }
        }
        elseif ($line -match "LICENSE DESCRIPTION:\s*(.*)") {
            $currentDescription = $matches[1].Trim()
            Write-Log "Found license description: $currentDescription" -Level Info
            
            if (-not $currentProduct) {
                # Fallback detection based on description
                if ($currentDescription -match "Office.*Professional Plus.*(2024|2021|2019)|Office.*(24|21|19).*RETAIL") {
                    $year = if ($matches[1]) { $matches[1] } else { $matches[2] }
                    if ($year -eq "24") { $year = "2024" }
                    elseif ($year -eq "21") { $year = "2021" }
                    elseif ($year -eq "19") { $year = "2019" }
                    $currentProduct = "ProPlus$year"
                }
                elseif ($currentDescription -match "Visio.*Professional.*(2024|2021|2019)") {
                    $currentProduct = "VisioPro$($matches[1])"
                }
                elseif ($currentDescription -match "Project.*Professional.*(2024|2021|2019)") {
                    $currentProduct = "ProjectPro$($matches[1])"
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
                
                # Extract version from product type
                $version = $OfficeVersion
                if ($currentProduct -match "(2024|2021|2019)") {
                    $version = $matches[1]
                }
                
                $product = @{
                    Type = $currentProduct
                    License = $licenseType
                    NeedsConversion = $needsConversion
                    Version = $version
                    PartialKey = $partialKey
                }
                
                $products += $product
                Write-Log "Found product: $currentProduct ($licenseType) - Version: $version - Key: $partialKey - Needs conversion: $needsConversion" -Level Info
                
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
        [string]$OfficeArchitecture,
        [string]$OfficeVersion
    )
    
    $allProducts = @()
    
    Write-Log "Retrieving all Office products for detailed status..." -Level Info
    $osppOutput = Get-InstalledProducts -ScriptPath $ScriptPath
    if ($osppOutput) {
        $osppProducts = Parse-OSPPOutput -Output $osppOutput -OfficeVersion $OfficeVersion
        if ($osppProducts.Count -gt 0) {
            Write-Log "Found $($osppProducts.Count) products via OSPP" -Level Success
            foreach ($product in $osppProducts) {
                Write-Log "Product: $($product.Type), License: $($product.License), Version: $($product.Version), Key: $($product.PartialKey), NeedsConversion: $($product.NeedsConversion)" -Level Info
            }
            $allProducts += $osppProducts
        }
    }
    
    if ($allProducts.Count -eq 0) {
        Write-Log "No products found via OSPP, checking registry..." -Level Warning
        $regProducts = Get-ProductsFromRegistry -OfficeArchitecture $OfficeArchitecture -OfficeVersion $OfficeVersion
        if ($regProducts.Count -gt 0) {
            Write-Log "Found $($regProducts.Count) products via registry" -Level Success
            foreach ($product in $regProducts) {
                Write-Log "Product: $($product.Type), License: $($product.License), Version: $($product.Version), NeedsConversion: $($product.NeedsConversion)" -Level Info
            }
            $allProducts += $regProducts
        }
    }
    
    $uniqueProducts = $allProducts | Sort-Object -Property Type -Unique
    return $uniqueProducts
}
#endregion

#region License Management
function Get-KMSKeys {
    return @{
        # Office 2024 KMS Keys
        "ProPlus2024" = "XM2V9-DN9HH-QB449-XDGKC-W2RMW"
        "VisioPro2024" = "JMMVY-XFNQC-KK4HK-9H7R3-WQQTV" 
        "ProjectPro2024" = "PD3TT-NTHQQ-VC7CY-P6KB6-BQ2C8"
        
        # Office 2021 KMS Keys  
        "ProPlus2021" = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
        "VisioPro2021" = "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4"
        "ProjectPro2021" = "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8"
        
        # Office 2019 KMS Keys
        "ProPlus2019" = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
        "VisioPro2019" = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
        "ProjectPro2019" = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
    }
}

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
    
    $result = Invoke-WithRetry -ScriptBlock {
        Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/unpkey:$PartialKey"
    } -MaxRetries 2 -OperationName "Remove retail license"
    
    if ($result.Success) {
        Write-Log "Retail license removed successfully" -Level Success
        
        # Reset license state
        Write-Log "Resetting license state with /rearm..." -Level Info
        $rearmResult = Invoke-WithRetry -ScriptBlock {
            Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/rearm"
        } -MaxRetries 2 -OperationName "Reset license state"
        
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
        [string]$OfficePath,
        [string]$OfficeVersion
    )
    
    $kmsKeys = Get-KMSKeys
    $key = $kmsKeys[$ProductType]
    
    if (-not $key) {
        Write-Log "No KMS key available for product type: $ProductType" -Level Error
        return $false
    }
    
    # Install KMS license files
    Write-Log "Checking and installing KMS Volume License files for $ProductType (Office $OfficeVersion)" -Level Info
    if (-not (Install-LicenseFiles -ScriptPath $ScriptPath -ProductType $ProductType -OfficePath $OfficePath -OfficeVersion $OfficeVersion)) {
        Write-Log "Failed to install KMS Volume License files for $ProductType" -Level Error
        return $false
    }
    
    Write-Log "Installing volume license for $ProductType with key: $key" -Level Info
    
    $result = Invoke-WithRetry -ScriptBlock {
        Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/inpkey:$key"
    } -MaxRetries 3 -OperationName "Install volume license"
    
    if ($result.Success) {
        Write-Log "Volume license installed successfully for $ProductType" -Level Success
        
        # Verify key installation
        Write-Log "Verifying key installation for $ProductType..." -Level Info
        $verifyResult = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/dstatus" -ReturnOutput
        if ($verifyResult.Success) {
            $keyEndChars = $key.Substring($key.Length - 5)
            if ($verifyResult.Output -match $keyEndChars) {
                Write-Log "Key verification successful: KMS key ending in $keyEndChars detected" -Level Success
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

function Test-KMSServer {
    param([string]$ServerAddress)
    
    if (-not $ServerAddress) {
        return $false
    }
    
    # Basic validation
    if ($ServerAddress -match "^[\w\.-]+(\:\d+)?$") {
        try {
            # Try to resolve the hostname
            $null = [System.Net.Dns]::GetHostAddresses($ServerAddress.Split(':')[0])
            return $true
        }
        catch {
            Write-Log "Cannot resolve hostname: $($ServerAddress.Split(':')[0])" -Level Warning
            return $false
        }
    }
    
    Write-Log "Invalid KMS server format: $ServerAddress" -Level Warning
    return $false
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
                do {
                    $kmsServer = Read-Host "Enter KMS server address (e.g., kms.example.com or 192.168.1.100:1688)"
                    if ($kmsServer -and (Test-KMSServer -ServerAddress $kmsServer)) {
                        Write-Log "Setting KMS server to: $kmsServer" -Level Info
                        
                        $result = Invoke-WithRetry -ScriptBlock {
                            Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/sethst:$kmsServer"
                        } -MaxRetries 2 -OperationName "Set KMS server"
                        
                        if ($result.Success) {
                            Write-Log "KMS server configured successfully" -Level Success
                            return $true
                        } else {
                            Write-Log "Failed to configure KMS server" -Level Error
                            return $false
                        }
                    }
                    elseif ($kmsServer) {
                        Write-Host "Invalid KMS server address. Please try again." -ForegroundColor Red
                    }
                } while ($kmsServer)
                return $true
            }
            "2" {
                Write-Log "Skipping KMS server configuration" -Level Info
                Write-Log "Warning: You may need to configure KMS server later for successful activation" -Level Warning
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
                
                $result = Invoke-WithRetry -ScriptBlock {
                    Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/act"
                } -MaxRetries 3 -DelaySeconds 5 -OperationName "Activate Office"
                
                if ($result.Success) {
                    Write-Log "Office activation attempted successfully" -Level Success
                    # Verify activation status
                    Start-Sleep -Seconds 3
                    $statusResult = Invoke-OSPPCommand -ScriptPath $ScriptPath -Arguments "/dstatusall" -ReturnOutput
                    if ($statusResult.Success) {
                        if ($statusResult.Output -match "LICENSE STATUS:.*LICENSED") {
                            Write-Log "Final verification: Office is fully activated (LICENSED)" -Level Success
                        } else {
                            Write-Log "Final verification: Office activation completed but may need more time to show LICENSED status." -Level Warning
                            Write-Log "You can check activation status later using: cscript `"$ScriptPath`" /dstatusall" -Level Info
                        }
                    } else {
                        Write-Log "Failed to verify activation status: $($statusResult.Error)" -Level Warning
                    }
                    return $true
                } else {
                    Write-Log "Office activation failed: $($result.Error)" -Level Error
                    return $false
                }
            }
            "2" {
                Write-Log "Skipping Office activation" -Level Info
                Write-Log "You can activate Office later using: cscript `"$ScriptPath`" /act" -Level Info
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
╔═══════════════════════════════════════════════════════════════════════╗
║        Office 2019/2021/2024 Retail to Volume License Converter       ║
║                    Version 0.7                                        ║
║                    Last update: 28-August-2025                        ║
║                    Author: Harry Hoang Le                             ║
║                                                                       ║
║  Supported Products:                                                  ║
║  • Microsoft Office Professional Plus 2019/2021/2024                  ║
║  • Microsoft Visio Professional 2019/2021/2024                        ║
║  • Microsoft Project Professional 2019/2021/2024                      ║
╚═══════════════════════════════════════════════════════════════════════╝
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
    Write-Log "Starting Office conversion process..." -Level Info
    
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
    
    $officeVersion = Get-OfficeVersion -OfficePath $officePath
    Write-Log "Detected Office version: $officeVersion" -Level Success
    
    if ($officeVersion -notin @("2019", "2021", "2024")) {
        Write-Log "Unsupported Office version: $officeVersion. This script supports Office 2019, 2021, and 2024 only." -Level Error
        return $false
    }
    
    $scriptPath = Join-Path $officePath "ospp.vbs"
    Write-Log "Using OSPP script: $scriptPath" -Level Info
    
    $products = Get-OfficeProductsAdvanced -ScriptPath $scriptPath -OfficeArchitecture $officeArch -OfficeVersion $officeVersion
    
    if ($products.Count -eq 0) {
        Write-Log "No Office $officeVersion products found for conversion" -Level Error
        return $false
    }
    
    Write-Log "Found $($products.Count) Office $officeVersion products" -Level Info
    
    foreach ($product in $products) {
        $status = if ($product.NeedsConversion) { "NEEDS CONVERSION" } else { "OK (Volume)" }
        Write-Log "Product: $($product.Type) - License: $($product.License) - Version: $($product.Version) - Key: $($product.PartialKey) - Status: $status" -Level Info
    }
    
    $productsToConvert = $products | Where-Object { $_.NeedsConversion -eq $true }
    
    if ($productsToConvert.Count -eq 0) {
        Write-Log "All detected products are already using Volume licensing" -Level Success
        Write-Log "Conversion process completed - no conversion needed" -Level Success
        return $true
    }
    
    Write-Log "Products requiring conversion: $($productsToConvert.Count)" -Level Info
    
    # Show conversion plan
    Write-Host "`nConversion Plan:" -ForegroundColor Yellow
    foreach ($product in $productsToConvert) {
        Write-Host "  → Convert $($product.Type) from Retail to Volume (Office $($product.Version))" -ForegroundColor White
    }
    
    Write-Host "`nDo you want to proceed with the conversion? (Y/N): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    
    if ($confirm -notmatch "^[Yy]") {
        Write-Log "Conversion cancelled by user" -Level Info
        return $false
    }
    
    Write-Log "Getting current license status for key removal..." -Level Info
    $currentStatus = Get-InstalledProducts -ScriptPath $scriptPath
    $currentProducts = Parse-OSPPOutput -Output $currentStatus -OfficeVersion $officeVersion
    
    $conversionSuccess = $true
    $convertedCount = 0
    
    foreach ($product in $productsToConvert) {
        Write-Log "Processing product: $($product.Type) (Office $($product.Version))" -Level Info
        
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
        
        if (Install-VolumeLicense -ScriptPath $scriptPath -ProductType $product.Type -OfficePath $officePath -OfficeVersion $product.Version) {
            $convertedCount++
            Write-Log "Successfully converted $($product.Type)" -Level Success
        }
        else {
            $conversionSuccess = $false
            Write-Log "Failed to convert $($product.Type)" -Level Error
        }
    }
    
    Write-Log "Conversion Summary: $convertedCount/$($productsToConvert.Count) products converted successfully" -Level Info
    
    if ($convertedCount -eq 0) {
        Write-Log "No products were successfully converted" -Level Error
        return $false
    }
    
    Write-Log "Verifying final status..." -Level Info
    Start-Sleep -Seconds 2
    
    $finalStatus = Get-InstalledProducts -ScriptPath $scriptPath
    $finalProducts = Parse-OSPPOutput -Output $finalStatus -OfficeVersion $officeVersion
    
    Write-Host "`nFinal Status Report:" -ForegroundColor Yellow
    foreach ($product in $finalProducts) {
        $status = if ($product.NeedsConversion) { "RETAIL" } else { "VOLUME" }
        $color = if ($product.NeedsConversion) { "Red" } else { "Green" }
        Write-Host "  $($product.Type) (Office $($product.Version)): $status - Key: $($product.PartialKey)" -ForegroundColor $color
    }
    
    $remainingRetail = $finalProducts | Where-Object { $_.NeedsConversion -eq $true }
    if ($remainingRetail.Count -gt 0) {
        Write-Log "Warning: $($remainingRetail.Count) products still show as Retail. Manual cleanup may be required." -Level Warning
        foreach ($retail in $remainingRetail) {
            Write-Log "Retail product remaining: $($retail.Type) with key: $($retail.PartialKey)" -Level Warning
        }
    }
    else {
        Write-Log "All products successfully converted to Volume licensing!" -Level Success
    }
    
    if ($convertedCount -gt 0) {
        Set-KMSServer -ScriptPath $scriptPath
        Start-OfficeActivation -ScriptPath $scriptPath
    }
    
    return $conversionSuccess
}

function Show-PostConversionInfo {
    Write-Host "`n" -NoNewline
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                   CONVERSION COMPLETED                    ║" -ForegroundColor Green  
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "Important Notes:" -ForegroundColor Yellow
    Write-Host "• Your Office products have been converted to Volume licensing" -ForegroundColor White
    Write-Host "• KMS activation requires connection to a KMS server" -ForegroundColor White
    Write-Host "• License files are installed from the Volume License package" -ForegroundColor White
    Write-Host "• For troubleshooting, check the log file saved in TEMP folder" -ForegroundColor White
    Write-Host ""
    Write-Host "Manual Commands (if needed):" -ForegroundColor Yellow
    Write-Host "• Check status: cscript `"C:\Program Files\Microsoft Office\Office16\ospp.vbs`" /dstatusall" -ForegroundColor Cyan
    Write-Host "• Set KMS server: cscript `"C:\Program Files\Microsoft Office\Office16\ospp.vbs`" /sethst:<server>" -ForegroundColor Cyan
    Write-Host "• Activate: cscript `"C:\Program Files\Microsoft Office\Office16\ospp.vbs`" /act" -ForegroundColor Cyan
    Write-Host ""
}

function Show-SupportedProducts {
    Write-Host "`nSupported Products and KMS Keys:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Office 2024:" -ForegroundColor Cyan
    Write-Host "  • ProPlus 2024: XM2V9-DN9HH-QB449-XDGKC-W2RMW" -ForegroundColor White
    Write-Host "  • Visio Pro 2024: JMMVY-XFNQC-KK4HK-9H7R3-WQQTV" -ForegroundColor White
    Write-Host "  • Project Pro 2024: PD3TT-NTHQQ-VC7CY-P6KB6-BQ2C8" -ForegroundColor White
    Write-Host ""
    Write-Host "Office 2021:" -ForegroundColor Cyan
    Write-Host "  • ProPlus 2021: FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH" -ForegroundColor White
    Write-Host "  • Visio Pro 2021: KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4" -ForegroundColor White
    Write-Host "  • Project Pro 2021: FTNWT-C6WBT-8HMGF-K9PRX-QV9H8" -ForegroundColor White
    Write-Host ""
    Write-Host "Office 2019:" -ForegroundColor Cyan
    Write-Host "  • ProPlus 2019: NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" -ForegroundColor White
    Write-Host "  • Visio Pro 2019: 9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" -ForegroundColor White
    Write-Host "  • Project Pro 2019: B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" -ForegroundColor White
    Write-Host ""
}

function Main {
    Show-Banner
    
    # Check for help parameter
    if ($args -contains "-Help" -or $args -contains "-h" -or $args -contains "/?") {
        Show-SupportedProducts
        Write-Host "Usage:" -ForegroundColor Yellow
        Write-Host "  .\script.ps1        - Run conversion process" -ForegroundColor White
        Write-Host "  .\script.ps1 -Debug - Run with debug output" -ForegroundColor White
        Write-Host "  .\script.ps1 -Help  - Show this help" -ForegroundColor White
        return
    }
    
    try {
        if (-not (Test-Prerequisites)) {
            Write-Log "Prerequisites not met. Exiting." -Level Error
            return
        }
        
        $success = Start-ConversionProcess
        
        if ($success) {
            Write-Log "Conversion process completed successfully!" -Level Success
            Show-PostConversionInfo
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

# Handle command line parameters
if ($args -contains "-Debug") {
    $script:DebugMode = $true
    Write-Log "Debug mode enabled" -Level Info
}

# Start main execution
Main
#endregion