# File: C:\Scripts\CheckCaptivePortal.ps1

# 1. Wait briefly for IP negotiation (DHCP) to finish after the event fires
Start-Sleep -Seconds 5

# 2. Load required assembly for the popup box
Add-Type -AssemblyName System.Windows.Forms

function Test-CaptivePortal {
    $TargetUri = "http://www.msftconnecttest.com/connecttest.txt"
    $ExpectedContent = "Microsoft Connect Test"

    try {
        $Response = Invoke-WebRequest -Uri $TargetUri -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        
        # If content matches Microsoft's text, we have real internet.
        if ($Response.Content -eq $ExpectedContent) {
            return $false
        }
        # If content differs, it's likely a login page.
        else {
            return $true
        }
    }
    catch {
        # Redirects (302) are the most common sign of a portal
        if ($_.Exception.Response.StatusCode -match "Found|Moved|Redirect") {
            return $true
        }
        return $false
    }
}

# 3. Execution Logic
# We check if we are connected to a non-internet network first to save processing time
$Profile = Get-NetConnectionProfile | Where-Object { $_.IPv4Connectivity -ne "Internet" } | Select-Object -First 1

if ($Profile) {
    if (Test-CaptivePortal) {
        # Show Popup
        $Result = [System.Windows.Forms.MessageBox]::Show(
            "Captive Portal Detected on: $($Profile.Name).`n`nPlease accept the EULA to proceed.",
            "Network Alert",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        # Launch Browser
        Start-Process "http://www.msftconnecttest.com/connecttest.txt"
    }
}