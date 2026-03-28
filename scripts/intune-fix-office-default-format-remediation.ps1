<#
.SYNOPSIS
    Detects and remediates incorrect default file formats in Office 365 applications (Excel, Word, PowerPoint).

.DESCRIPTION
    This script combines detection and remediation logic for an Intune Proactive Remediation package.
    It checks whether Excel, Word, and PowerPoint are configured to use their native Office file formats
    (XLSX, DOCX, PPTX) rather than open document formats (ODS, ODT, ODP).

    When deployed via Intune Proactive Remediation:
      - The DETECTION phase checks the HKCU registry keys for each Office application's DefaultFormat value.
        If any value is incorrect, the script exits with code 1 to trigger remediation.
      - The REMEDIATION phase sets the correct DefaultFormat registry values for all three applications:
          Excel  -> 51   (XLSX)
          Word   -> DOCX (string)
          PowerPoint -> 1 (PPTX)

    This script should be split into two separate .ps1 files for Intune deployment:
      - Detect-OfficeDefaultFormat.ps1  (Detection script)
      - Remediate-OfficeDefaultFormat.ps1 (Remediation script)

    Both scripts are included below, clearly separated by region blocks for easy extraction.

.NOTES
    Author:      Souhaiel Morhag
    Company:     MSEndpoint.com
    Blog:        https://msendpoint.com
    Academy:     https://app.msendpoint.com/academy
    LinkedIn:    https://linkedin.com/in/souhaiel-morhag
    GitHub:      https://github.com/Msendpoint
    License:     MIT

.EXAMPLE
    # Run detection locally to check current Office default format settings:
    .\Detect-OfficeDefaultFormat.ps1

    # Run remediation locally to correct Office default format settings:
    .\Remediate-OfficeDefaultFormat.ps1

    # In Intune, upload each script separately under:
    # Devices > Scripts and remediations > Create > Proactive remediation
#>

#region -------------------------------------------------------
# DETECTION SCRIPT: Detect-OfficeDefaultFormat.ps1
# Deploy this block as your Intune Proactive Remediation DETECTION script.
# Exit 0 = compliant | Exit 1 = non-compliant (triggers remediation)
#---------------------------------------------------------------

[CmdletBinding()]
param()

# Collect names of applications with incorrect default format settings
$incorrectSettings = @()

try {
    # --- Excel: DefaultFormat 51 = XLSX ---
    $excelKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    $excelValue = Get-ItemProperty -Path $excelKey -Name 'DefaultFormat' -ErrorAction SilentlyContinue

    if ($null -eq $excelValue -or $excelValue.DefaultFormat -ne 51) {
        $incorrectSettings += 'Excel (expected 51 for XLSX)'
    }

    # --- Word: DefaultFormat 'DOCX' (string) = DOCX ---
    $wordKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Options'
    $wordValue = Get-ItemProperty -Path $wordKey -Name 'DefaultFormat' -ErrorAction SilentlyContinue

    if ($null -eq $wordValue -or $wordValue.DefaultFormat -ne 'DOCX') {
        $incorrectSettings += 'Word (expected DOCX string)'
    }

    # --- PowerPoint: DefaultFormat 1 = PPTX ---
    $pptKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options'
    $pptValue = Get-ItemProperty -Path $pptKey -Name 'DefaultFormat' -ErrorAction SilentlyContinue

    if ($null -eq $pptValue -or $pptValue.DefaultFormat -ne 1) {
        $incorrectSettings += 'PowerPoint (expected 1 for PPTX)'
    }

    if ($incorrectSettings.Count -gt 0) {
        Write-Host "Non-compliant: Incorrect default format detected for: $($incorrectSettings -join ', ')"
        Exit 1
    }
    else {
        Write-Host 'Compliant: All Office applications are using the correct native default file format.'
        Exit 0
    }
}
catch {
    Write-Host "Detection error: $($_.Exception.Message)"
    Exit 1
}

#endregion


#region -------------------------------------------------------
# REMEDIATION SCRIPT: Remediate-OfficeDefaultFormat.ps1
# Deploy this block as your Intune Proactive Remediation REMEDIATION script.
# This script corrects the DefaultFormat registry values for Excel, Word, and PowerPoint.
#---------------------------------------------------------------

<#
[CmdletBinding()]
param()

try {
    # --- Set Excel default format to XLSX (value: 51) ---
    $excelKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    if (-not (Test-Path $excelKey)) {
        New-Item -Path $excelKey -Force | Out-Null
    }
    Set-ItemProperty -Path $excelKey -Name 'DefaultFormat' -Value 51 -Type DWord
    Write-Host 'Excel: DefaultFormat set to 51 (XLSX).'

    # --- Set Word default format to DOCX (value: 'DOCX' string) ---
    $wordKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Options'
    if (-not (Test-Path $wordKey)) {
        New-Item -Path $wordKey -Force | Out-Null
    }
    Set-ItemProperty -Path $wordKey -Name 'DefaultFormat' -Value 'DOCX' -Type String
    Write-Host 'Word: DefaultFormat set to DOCX.'

    # --- Set PowerPoint default format to PPTX (value: 1) ---
    $pptKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options'
    if (-not (Test-Path $pptKey)) {
        New-Item -Path $pptKey -Force | Out-Null
    }
    Set-ItemProperty -Path $pptKey -Name 'DefaultFormat' -Value 1 -Type DWord
    Write-Host 'PowerPoint: DefaultFormat set to 1 (PPTX).'

    Write-Host 'Remediation complete: Default formats corrected for Excel, Word, and PowerPoint.'
    Exit 0
}
catch {
    Write-Host "Remediation error: $($_.Exception.Message)"
    Exit 1
}
#>

#endregion
