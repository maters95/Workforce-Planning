Add-Type -AssemblyName System.Windows.Forms

# ─────────────────────────────────────────────────────────────────
# TEST MODE
#   $true  → opens batch-driving-record.html in your browser and
#            sends keystrokes there (window title: "Test Terminal")
#   $false → targets the real DRIVES Java terminal (window title: "DRIVES")
# ─────────────────────────────────────────────────────────────────
$testMode       = $true
$testHtmlPath   = "file:///C:/Users/zmasters1/Downloads/batch-driving-record.html"
$drivesWindow   = if ($testMode) { "Test Terminal" } else { "DRIVES" }

# ─────────────────────────────────────────────────────────────────
# 1. FIND THE BATCH EMAIL IN OUTLOOK INBOX
# ─────────────────────────────────────────────────────────────────
$outlook = New-Object -ComObject Outlook.Application
$ns      = $outlook.GetNamespace("MAPI")
$inbox   = $ns.GetDefaultFolder(6)   # 6 = olFolderInbox

$batchEmail = $null
foreach ($msg in $inbox.Items) {
    if ($msg.UnRead -and $msg.Subject -like "*DRIVES Batching -*") {
        $batchEmail = $msg
        break
    }
}

if ($null -eq $batchEmail) {
    [System.Windows.Forms.MessageBox]::Show(
        "No unread batch email found in your inbox.`nRun the cloud flow first.",
        "DRIVES Batch Entry", "OK", "Warning") | Out-Null
    exit
}

# ─────────────────────────────────────────────────────────────────
# 2. EXTRACT NUMBERS FROM THE <pre>...</pre> BLOCK
# ─────────────────────────────────────────────────────────────────
$htmlBody = $batchEmail.HTMLBody
if ($htmlBody -notmatch '(?s)<pre[^>]*>(.*?)</pre>') {
    [System.Windows.Forms.MessageBox]::Show(
        "No number list found in the email body.",
        "DRIVES Batch Entry", "OK", "Warning") | Out-Null
    exit
}

$preRaw = $matches[1]
# Strip any inline HTML tags, decode entities, split on newlines
$preRaw = $preRaw -replace '<[^>]+>', ''
$preRaw = [System.Net.WebUtility]::HtmlDecode($preRaw)
$numbers = $preRaw -split "`r?`n" |
           ForEach-Object { $_.Trim() } |
           Where-Object   { $_ -ne '' }

if ($numbers.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "The batch email was found but contained no numbers.",
        "DRIVES Batch Entry", "OK", "Warning") | Out-Null
    exit
}

# ─────────────────────────────────────────────────────────────────
# 3. CONFIRM WITH USER
# ─────────────────────────────────────────────────────────────────
$result = [System.Windows.Forms.MessageBox]::Show(
    "Ready to enter $($numbers.Count) number(s) into DRIVES.`n`nMake sure:`n  1. DRIVES Java terminal is open`n  2. You are on the Main Menu (screen 0)`n`nClick OK to begin.",
    "DRIVES Batch Entry", "OKCancel", "Information")

if ($result -eq "Cancel") { exit }

# ─────────────────────────────────────────────────────────────────
# 4. FOCUS THE TARGET WINDOW
# ─────────────────────────────────────────────────────────────────
$wsh = New-Object -ComObject WScript.Shell

if ($testMode) {
    # Open the HTML test page in the default browser then wait for it to load
    Start-Process $testHtmlPath
    Start-Sleep -Milliseconds 2000
}

$wsh.AppActivate($drivesWindow) | Out-Null
Start-Sleep -Milliseconds 600

function Send-Keys($keys, [int]$waitMs = 400) {
    $wsh.AppActivate($drivesWindow) | Out-Null
    $wsh.SendKeys($keys)
    Start-Sleep -Milliseconds $waitMs
}

# ─────────────────────────────────────────────────────────────────
# 5. NAVIGATE TO THE BATCH DRIVING RECORD REQUEST SCREEN
#    3 → Enter → 5 → Enter → 14 → Enter → 7 → Enter
# ─────────────────────────────────────────────────────────────────
Send-Keys "3~"    # Main Menu → Record Services
Send-Keys "5~"    # Record Services → Batch Operations
Send-Keys "14~"   # Batch Operations → Batch Driving Record
Send-Keys "7~"    # Batch Driving Record → Batch Driving Record Request

# ─────────────────────────────────────────────────────────────────
# 6. FILL THE HEADER: Y (auto) → N (auto) → Y (auto) → Enter
# ─────────────────────────────────────────────────────────────────
Send-Keys "Y"  200
Send-Keys "N"  200
Send-Keys "Y"  200
Send-Keys "~"  500   # Enter → opens email selector screen

# ─────────────────────────────────────────────────────────────────
# 7. SELECT EMAIL: Tab → S → F5 → Tab → Tab (lands on grid)
# ─────────────────────────────────────────────────────────────────
Send-Keys "{TAB}"       200   # move to Sel field
Send-Keys "S"           200   # select the CAU email row
Send-Keys "{F5}"        500   # return to batch form
Send-Keys "{TAB}{TAB}"  200   # land on first grid cell

# ─────────────────────────────────────────────────────────────────
# 8. TYPE NUMBERS INTO THE GRID
#    90 cells per screen (15 rows × 6 cols).
#    8-char → just type (auto-advances).
#    Everything else → type + Tab.
# ─────────────────────────────────────────────────────────────────
$maxPerScreen = 90
$offset       = 0
$total        = $numbers.Count

while ($offset -lt $total) {

    $batch = 0
    while ($batch -lt $maxPerScreen -and ($offset + $batch) -lt $total) {
        $num = $numbers[$offset + $batch]

        if ($num.Length -eq 8) {
            $wsh.SendKeys($num)          # auto-advances after 8 chars
        } else {
            $wsh.SendKeys($num + "{TAB}") # manual advance for 7-char / alphanumeric
        }
        Start-Sleep -Milliseconds 60
        $batch++
    }

    # Submit this screenful
    Send-Keys "~" 1200   # Enter + wait for screen to process

    $offset += $batch

    if ($offset -lt $total) {
        # Overflow: navigate back for the next screen
        Send-Keys "{F5}"       600
        Send-Keys "7~"         600
        Send-Keys "Y"          200
        Send-Keys "N"          200
        Send-Keys "Y"          200
        Send-Keys "~"          600
        Send-Keys "{TAB}"      200
        Send-Keys "S"          200
        Send-Keys "{F5}"       600
        Send-Keys "{TAB}{TAB}" 200
    }
}

# ─────────────────────────────────────────────────────────────────
# 9. MARK EMAIL AS READ AND FINISH
# ─────────────────────────────────────────────────────────────────
$batchEmail.UnRead = $false

[System.Windows.Forms.MessageBox]::Show(
    "All $total number(s) submitted to DRIVES successfully.",
    "DRIVES Batch Entry Complete", "OK", "Information") | Out-Null
