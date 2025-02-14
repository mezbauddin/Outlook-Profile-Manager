<#
.SYNOPSIS
  Create/update an Outlook profile and set per-account default signatures directly in the profile registry.

.DESCRIPTION
  1) Closes Outlook.
  2) Creates or imports a new profile.
  3) Restores signature files from backup.
  4) Ensures .htm, .rtf, and .txt signature files exist.
  5) Enumerates account subkeys in the new profile and sets "New Signature" and "Reply-Forward Signature".
  6) Disables the profile prompt so that the new profile is always used.
#>

# -----------------------
# 0) Variables
# -----------------------
$officeVersion  = "16.0"                # 16.0 = Outlook 2016/2019/365
$newProfileName = "MyNewProfile"
$prfFile        = "C:\Deployment\CorporateExchangeProfile.prf"

# Registry and signature paths
$regOutlook    = "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook"
$profileBase   = Join-Path $regOutlook "Profiles"
$signatureDir  = "$($env:APPDATA)\Microsoft\Signatures"
$signatureBackup = "C:\Temp\SignatureBackup\$($env:USERNAME)"

# -----------------------
# 1) Close Outlook
# -----------------------
Write-Host "Closing Outlook if running..."
$processes = Get-Process outlook -ErrorAction SilentlyContinue
if ($processes) {
    $processes | Stop-Process -Force
    Start-Sleep -Seconds 3
}

# -----------------------
# 2) Create/Import the New Profile
# -----------------------
Write-Host "Creating or updating Outlook profile: $newProfileName"

if (Test-Path $prfFile) {
    Write-Host "Importing PRF file..."
    Start-Process "outlook.exe" -ArgumentList "/importprf `"$prfFile`"" -Wait
    Start-Sleep -Seconds 5
} else {
    Write-Host "No PRF found at $prfFile. Using minimal registry approach..."
    $newProfileKey = Join-Path $profileBase $newProfileName
    if (!(Test-Path $newProfileKey)) {
        New-Item -Path $newProfileKey -Force | Out-Null
    }
    # Note: In a real scenario, you would populate this profile's subkeys with mail server settings.
}

# Set the new profile as default and disable the profile prompt.
Set-ItemProperty -Path $regOutlook -Name "DefaultProfile" -Value $newProfileName
Set-ItemProperty -Path $regOutlook -Name "PromptForProfile" -Value 0

# -----------------------
# 3) Restore Signatures
# -----------------------
if (Test-Path $signatureBackup) {
    Write-Host "Restoring signatures from $signatureBackup to $signatureDir"
    if (!(Test-Path $signatureDir)) {
        New-Item -Path $signatureDir -ItemType Directory | Out-Null
    }
    Copy-Item -Path "$signatureBackup\*" -Destination $signatureDir -Recurse -Force
} else {
    Write-Host "No signature backup found; skipping restore."
}

# -----------------------
# 4) Ensure .htm, .rtf, and .txt files exist for each signature
# -----------------------
$sigHtmFiles = Get-ChildItem -Path $signatureDir -Filter "*.htm" -File -ErrorAction SilentlyContinue

if ($sigHtmFiles) {
    foreach ($htmFile in $sigHtmFiles) {
        $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($htmFile.Name)
        $sigRtfPath = Join-Path $signatureDir ($baseName + ".rtf")
        $sigTxtPath = Join-Path $signatureDir ($baseName + ".txt")
        
        if (!(Test-Path $sigRtfPath)) {
            Write-Host "Creating placeholder RTF for signature '$baseName'"
            "This is the RTF version of $baseName" | Out-File $sigRtfPath -Encoding ASCII
        }
        if (!(Test-Path $sigTxtPath)) {
            Write-Host "Creating placeholder TXT for signature '$baseName'"
            "This is the TXT version of $baseName" | Out-File $sigTxtPath -Encoding ASCII
        }
    }
}

# Pick the FIRST .htm signature as our default
$firstSig = $sigHtmFiles | Select-Object -First 1
if (!$firstSig) {
    Write-Host "No .htm signature found in $signatureDir; can't set default signature."
} else {
    $defaultSignatureName = [System.IO.Path]::GetFileNameWithoutExtension($firstSig.Name)
    Write-Host "Will set default signature to '$defaultSignatureName'"
}

# -----------------------
# 5) Set "New Signature" and "Reply-Forward Signature" in the Profile's Account Subkeys
# -----------------------
if ($defaultSignatureName) {
    # Outlook stores account data under the fixed GUID subkey "9375CFF0413111d3B88A00104B2A6676"
    $accountsPath = Join-Path (Join-Path $profileBase $newProfileName) "9375CFF0413111d3B88A00104B2A6676"

    if (Test-Path $accountsPath) {
        # Replace '-Directory' with a filter to select containers
        $accountKeys = Get-ChildItem -Path $accountsPath | Where-Object { $_.PSIsContainer }
        foreach ($acctKey in $accountKeys) {
            $acctPath = $acctKey.PSPath

            # Retrieve the "Account Name" property if it exists
            $acctName = (Get-ItemProperty -Path $acctPath -ErrorAction SilentlyContinue)."Account Name"
            if (!$acctName) {
                continue
            }

            Write-Host "Setting default signature for account '$acctName' in subkey '$($acctKey.Name)'"

            # Set the properties for default signatures
            New-ItemProperty -Path $acctPath -Name "New Signature" -Value $defaultSignatureName -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $acctPath -Name "Reply-Forward Signature" -Value $defaultSignatureName -PropertyType String -Force | Out-Null
        }
    } else {
        Write-Host "WARNING: Could not find $accountsPath. The profile might be incomplete."
    }
}

# -----------------------
# 6) Launch Outlook
# -----------------------
Write-Host "Launching Outlook with profile '$newProfileName'."
Start-Process "outlook.exe"

Write-Host "`nDone! The signature files are restored and default signature settings have been applied per account."
Write-Host "If the signature still does not appear, verify Outlook’s signature settings via File → Options → Mail → Signatures."
