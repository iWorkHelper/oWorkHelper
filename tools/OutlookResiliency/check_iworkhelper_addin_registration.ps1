#Requires -Version 3.0
<#
.SYNOPSIS
    Check iWorkHelper Outlook add-in registration and deployment status.

.DESCRIPTION
    This script scans Windows registry for iWorkHelper add-in configuration,
    collecting LoadBehavior, Manifest path, deployment details, etc.
    Read-only, no modifications.

.PARAMETER OutputFile
    Output file path for diagnosis report. If not specified, output to console.

.EXAMPLE
    .\check_iworkhelper_addin_registration.ps1
    .\check_iworkhelper_addin_registration.ps1 -OutputFile "C:\diagnosis.txt"

.NOTES
    - No admin rights required (read-only)
    - No registry modifications
    - No sensitive information in output
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = $null
)

function Write-DiagnosisLine {
    param([string]$Line)
    if ($OutputFile) {
        Add-Content -Path $OutputFile -Value $Line -Encoding UTF8
    }
    else {
        Write-Host $Line
    }
}

function Write-DiagnosisHeader {
    param([string]$Header)
    $line = "=" * 80
    Write-DiagnosisLine $line
    Write-DiagnosisLine $Header
    Write-DiagnosisLine $line
}

function Check-RegistryPath {
    param([string]$Path, [string]$Description)

    Write-DiagnosisLine ""
    Write-DiagnosisLine "[$Description]"
    Write-DiagnosisLine "Path: $Path"

    try {
        if (Test-Path -Path "Registry::$Path") {
            $subItems = Get-ChildItem -Path "Registry::$Path" -ErrorAction SilentlyContinue
            if (-not $subItems -or $subItems.Count -eq 0) {
                Write-DiagnosisLine "Status: Path exists but empty"
                return
            }

            Write-DiagnosisLine "Status: Found items:"
            foreach ($subItem in $subItems) {
                $name = $subItem.PSChildName
                if ($name -match "(iWorkHelper|eWorkHelper|oWorkHelper|ThisAddIn)") {
                    Write-DiagnosisLine ""
                    Write-DiagnosisLine "  Item: $name"

                    try {
                        $props = Get-ItemProperty -Path "Registry::$Path\$name" -ErrorAction SilentlyContinue

                        if ($props.Manifest) {
                            Write-DiagnosisLine "    Manifest: $($props.Manifest)"
                        }
                        if ($props.LoadBehavior) {
                            $lbValue = $props.LoadBehavior
                            $lbDesc = switch ($lbValue) {
                                3 { "Startup Load (Default)" }
                                9 { "Disabled (Slow Load)" }
                                16 { "Load on Demand" }
                                2 { "Depends on Setting" }
                                8 { "Load Failed" }
                                0 { "No Auto Load" }
                                default { "Unknown" }
                            }
                            Write-DiagnosisLine "    LoadBehavior: $lbValue ($lbDesc)"
                        }
                        if ($props.FriendlyName) {
                            Write-DiagnosisLine "    FriendlyName: $($props.FriendlyName)"
                        }
                        if ($props.Description) {
                            Write-DiagnosisLine "    Description: $($props.Description)"
                        }
                        if ($props.Location) {
                            Write-DiagnosisLine "    Location: $($props.Location)"
                        }
                    }
                    catch {
                        Write-DiagnosisLine "    [Unable to read properties]"
                    }
                }
            }
        }
        else {
            Write-DiagnosisLine "Status: Path not found"
        }
    }
    catch {
        Write-DiagnosisLine "Status: Access failed - $($_.Exception.Message)"
    }
}

function Check-ResiliencyStatus {
    param([string]$Root)

    $resPath = "$Root\Software\Microsoft\Office\16.0\Outlook\Resiliency"
    Write-DiagnosisLine ""
    Write-DiagnosisLine "Checking Resiliency Status..."
    Write-DiagnosisLine "Path: $resPath"

    try {
        if (Test-Path -Path "Registry::$resPath") {
            # Check LoadBehavior
            $lbPath = "$resPath\LoadBehavior"
            if (Test-Path -Path "Registry::$lbPath") {
                $lbItems = Get-ChildItem -Path "Registry::$lbPath" -ErrorAction SilentlyContinue
                foreach ($item in $lbItems) {
                    $name = $item.PSChildName
                    if ($name -match "(iWorkHelper|eWorkHelper|oWorkHelper|ThisAddIn)") {
                        $value = Get-ItemProperty -Path "Registry::$lbPath\$name" -Name "(Default)" -ErrorAction SilentlyContinue
                        Write-DiagnosisLine "  LoadBehavior - $name : $($value.'(Default)')"
                    }
                }
            }

            # Check DisabledItems
            $diPath = "$resPath\DisabledItems"
            if (Test-Path -Path "Registry::$diPath") {
                $diItems = Get-ItemProperty -Path "Registry::$diPath" -ErrorAction SilentlyContinue
                $found = $false
                foreach ($prop in $diItems.PSObject.Properties) {
                    if ($prop.Value -match "(iWorkHelper|eWorkHelper|oWorkHelper|ThisAddIn)") {
                        Write-DiagnosisLine "  DisabledItems found: $($prop.Value)"
                        $found = $true
                    }
                }
                if (-not $found) {
                    Write-DiagnosisLine "  DisabledItems: None"
                }
            }

            # Check CrashingAddinList
            $calPath = "$resPath\CrashingAddinList"
            if (Test-Path -Path "Registry::$calPath") {
                $calItems = Get-ItemProperty -Path "Registry::$calPath" -ErrorAction SilentlyContinue
                $found = $false
                foreach ($prop in $calItems.PSObject.Properties) {
                    if ($prop.Value -match "(iWorkHelper|eWorkHelper|oWorkHelper|ThisAddIn)") {
                        Write-DiagnosisLine "  CrashingAddinList found: $($prop.Value)"
                        $found = $true
                    }
                }
                if (-not $found) {
                    Write-DiagnosisLine "  CrashingAddinList: None"
                }
            }
        }
        else {
            Write-DiagnosisLine "Resiliency path not found (first run?)"
        }
    }
    catch {
        Write-DiagnosisLine "Resiliency check failed: $($_.Exception.Message)"
    }
}

function Check-InstalledVersions {
    Write-DiagnosisHeader "Checking Installed iWorkHelper Versions"

    $paths = @(
        "C:\Users\*\AppData\Local\Packages\iWorkhelper*",
        "C:\Users\*\AppData\Local\iWorkhelper*",
        "C:\Program Files\iWorkhelper*",
        "C:\Program Files (x86)\iWorkhelper*"
    )

    $found = $false
    foreach ($pattern in $paths) {
        $results = Get-Item -Path $pattern -ErrorAction SilentlyContinue
        if ($results) {
            $found = $true
            foreach ($result in $results) {
                Write-DiagnosisLine ""
                Write-DiagnosisLine "Found: $($result.FullName)"

                $dllFiles = Get-ChildItem -Path $result.FullName -Filter "iWorkhelper.dll" -Recurse -ErrorAction SilentlyContinue
                foreach ($dll in $dllFiles) {
                    $version = (Get-Item $dll.FullName).VersionInfo.FileVersion
                    Write-DiagnosisLine "  DLL: $($dll.FullName)"
                    Write-DiagnosisLine "  Version: $version"
                }
            }
        }
    }

    if (-not $found) {
        Write-DiagnosisLine "No iWorkHelper installation found"
    }
}

# Main execution
try {
    if ($OutputFile) {
        # Create output directory if not exists
        $outputDir = Split-Path -Path $OutputFile
        if ($outputDir -and -not (Test-Path -Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        "Diagnosis Report: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "Diagnosis started, output to: $OutputFile"
    }
    else {
        Write-Host "Diagnosis started..."
    }

    Write-DiagnosisHeader "iWorkHelper Outlook Add-in Diagnosis"
    Write-DiagnosisLine "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-DiagnosisLine "OS: $(Get-WmiObject Win32_OperatingSystem | Select-Object -ExpandProperty Caption)"
    Write-DiagnosisLine "PowerShell: $($PSVersionTable.PSVersion)"

    # Check HKCU
    # The iWorkHelper installer writes Software\Microsoft\Office\<Host>\Addins\<AddinId>
    # (no Office version segment), so both the version-agnostic and the Office 16.0 paths
    # must be probed; otherwise a successful install looks like "not found".
    Write-DiagnosisHeader "User Registry (HKEY_CURRENT_USER)"
    Check-RegistryPath "HKCU\Software\Microsoft\Office\Outlook\Addins" "HKCU Addins (version-agnostic, written by iWorkHelper installer)"
    Check-RegistryPath "HKCU\Software\WOW6432Node\Microsoft\Office\Outlook\Addins" "HKCU WOW6432Node Addins (version-agnostic, 32-bit view)"
    Check-RegistryPath "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins" "HKCU Addins (Office 16.0)"
    Check-RegistryPath "HKCU\Software\WOW6432Node\Microsoft\Office\16.0\Outlook\Addins" "HKCU WOW6432Node Addins (Office 16.0, 32-bit view)"
    Check-ResiliencyStatus "HKCU"

    # Check HKLM
    Write-DiagnosisHeader "Machine Registry (HKEY_LOCAL_MACHINE)"
    Check-RegistryPath "HKLM\Software\Microsoft\Office\Outlook\Addins" "HKLM Addins (version-agnostic, written by iWorkHelper installer)"
    Check-RegistryPath "HKLM\Software\WOW6432Node\Microsoft\Office\Outlook\Addins" "HKLM WOW6432Node Addins (version-agnostic, 32-bit view)"
    Check-RegistryPath "HKLM\Software\Microsoft\Office\16.0\Outlook\Addins" "HKLM Addins (Office 16.0)"
    Check-RegistryPath "HKLM\Software\WOW6432Node\Microsoft\Office\16.0\Outlook\Addins" "HKLM WOW6432Node Addins (Office 16.0, 32-bit view)"
    Check-ResiliencyStatus "HKLM"

    # Check versions
    Check-InstalledVersions

    # Summary
    Write-DiagnosisHeader "Diagnosis Complete"
    Write-DiagnosisLine "Please check:"
    Write-DiagnosisLine "1. Which version of DLL is Outlook loading?"
    Write-DiagnosisLine "2. What is the LoadBehavior value (3=startup, 9=disabled, 16=on-demand)?"
    Write-DiagnosisLine "3. Are there multiple versions?"
    Write-DiagnosisLine "4. Is it in DisabledItems or CrashingAddinList?"
    Write-DiagnosisLine "5. The iWorkHelper installer registers add-ins under Software\Microsoft\Office\<Host>\Addins"
    Write-DiagnosisLine "   (no version segment) with key names 'oWorkhelper' / 'eWorkhelper'; Office 16.0 paths are"
    Write-DiagnosisLine "   listed for reference only, so a match there is not required for a successful install."
    Write-DiagnosisLine ""

    if ($OutputFile) {
        Write-Host "Diagnosis complete, report saved to: $OutputFile"
    }
}
catch {
    Write-Error "Script error: $($_.Exception.Message)"
    exit 1
}
