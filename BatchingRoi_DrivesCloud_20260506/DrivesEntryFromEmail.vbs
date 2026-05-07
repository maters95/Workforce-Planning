Option Explicit

' ─────────────────────────────────────────────────────────────────
' TEST MODE
'   True  → opens batch-driving-record.html in browser
'           (targets window title: "Test Terminal")
'   False → targets real DRIVES Java terminal (title: "DRIVES")
' ─────────────────────────────────────────────────────────────────
Const TEST_MODE    = True
Const TEST_HTML    = "file:///C:/Users/zmasters1/Downloads/batch-driving-record.html"
Const DRIVES_TITLE = "DRIVES"
Const TEST_TITLE   = "Test Terminal"

Dim targetWindow
If TEST_MODE Then targetWindow = TEST_TITLE Else targetWindow = DRIVES_TITLE

' ─────────────────────────────────────────────────────────────────
' 1. FIND THE BATCH EMAIL IN OUTLOOK INBOX
' ─────────────────────────────────────────────────────────────────
Dim outlook, ns, inbox, msg, batchEmail
Set outlook = CreateObject("Outlook.Application")
Set ns      = outlook.GetNamespace("MAPI")
Set inbox   = ns.GetDefaultFolder(6)   ' 6 = olFolderInbox

inbox.Items.Sort "[ReceivedTime]", True  ' newest first

Set batchEmail = Nothing
For Each msg In inbox.Items
    If msg.UnRead And InStr(msg.Subject, "DRIVES Batching -") > 0 Then
        Set batchEmail = msg
        Exit For
    End If
Next

If batchEmail Is Nothing Then
    MsgBox "No unread batch email found in your inbox." & vbCrLf & _
           "Run the cloud flow first.", _
           vbExclamation, "DRIVES Batch Entry"
    WScript.Quit
End If

' ─────────────────────────────────────────────────────────────────
' 2. EXTRACT NUMBERS FROM THE <pre>...</pre> BLOCK
' ─────────────────────────────────────────────────────────────────
Dim htmlBody, re, matches, preRaw
htmlBody = batchEmail.HTMLBody

Set re          = New RegExp
re.Pattern      = "<pre[^>]*>([\s\S]*?)</pre>"
re.IgnoreCase   = True
Set matches = re.Execute(htmlBody)

If matches.Count = 0 Then
    MsgBox "No number list found in the email body.", _
           vbExclamation, "DRIVES Batch Entry"
    WScript.Quit
End If

preRaw = matches(0).SubMatches(0)

' Strip HTML tags
Set re        = New RegExp
re.Pattern    = "<[^>]+>"
re.Global     = True
preRaw = re.Replace(preRaw, "")

' Decode common HTML entities
preRaw = Replace(preRaw, "&amp;",  "&")
preRaw = Replace(preRaw, "&lt;",   "<")
preRaw = Replace(preRaw, "&gt;",   ">")
preRaw = Replace(preRaw, "&nbsp;", " ")

' Normalise line endings then split
preRaw = Replace(preRaw, vbCrLf, vbLf)
preRaw = Replace(preRaw, vbCr,   vbLf)

Dim rawLines, ln
rawLines = Split(preRaw, vbLf)

Dim numbers()
Dim numCount : numCount = 0
ReDim numbers(UBound(rawLines))

For Each ln In rawLines
    ln = Trim(ln)
    If ln <> "" Then
        numbers(numCount) = ln
        numCount = numCount + 1
    End If
Next

If numCount = 0 Then
    MsgBox "The batch email was found but contained no numbers.", _
           vbExclamation, "DRIVES Batch Entry"
    WScript.Quit
End If
ReDim Preserve numbers(numCount - 1)

' ─────────────────────────────────────────────────────────────────
' 3. CONFIRM WITH USER
' ─────────────────────────────────────────────────────────────────
Dim confirm
confirm = MsgBox( _
    "Ready to enter " & numCount & " number(s) into DRIVES." & vbCrLf & vbCrLf & _
    "Make sure:" & vbCrLf & _
    "  1. DRIVES Java terminal is open" & vbCrLf & _
    "  2. You are on the Main Menu (screen 0)" & vbCrLf & vbCrLf & _
    "Click OK to begin.", _
    vbOKCancel + vbInformation, "DRIVES Batch Entry")

If confirm = vbCancel Then WScript.Quit

' ─────────────────────────────────────────────────────────────────
' 4. FOCUS THE TARGET WINDOW
' ─────────────────────────────────────────────────────────────────
Dim wsh
Set wsh = CreateObject("WScript.Shell")

If TEST_MODE Then
    wsh.Run TEST_HTML
    WScript.Sleep 2000
End If

wsh.AppActivate targetWindow
WScript.Sleep 600

' Helper sub: activate window, send keys, wait
Sub SK(keys, waitMs)
    wsh.AppActivate targetWindow
    wsh.SendKeys keys
    WScript.Sleep waitMs
End Sub

' ─────────────────────────────────────────────────────────────────
' 5. NAVIGATE TO BATCH DRIVING RECORD REQUEST SCREEN
'    3 → Enter → 5 → Enter → 14 → Enter → 7 → Enter
' ─────────────────────────────────────────────────────────────────
SK "3~",  400
SK "5~",  400
SK "14~", 400
SK "7~",  400

' ─────────────────────────────────────────────────────────────────
' 6. HEADER: Y (auto) → N (auto) → Y (auto) → Enter
' ─────────────────────────────────────────────────────────────────
SK "Y", 200
SK "N", 200
SK "Y", 200
SK "~", 500

' ─────────────────────────────────────────────────────────────────
' 7. EMAIL SELECT: Tab → S → F5 → Tab → Tab (land on grid)
' ─────────────────────────────────────────────────────────────────
SK "{TAB}",      200
SK "S",          200
SK "{F5}",       500
SK "{TAB}{TAB}", 200

' ─────────────────────────────────────────────────────────────────
' 8. TYPE NUMBERS INTO THE GRID
'    90 cells per screen (15 rows x 6 cols).
'    8-char → just type (auto-advances).
'    Everything else → type + Tab.
' ─────────────────────────────────────────────────────────────────
Dim maxPerScreen, offset, total, batch, num
maxPerScreen = 90
offset       = 0
total        = numCount

Do While offset < total

    batch = 0
    Do While batch < maxPerScreen And (offset + batch) < total
        num = numbers(offset + batch)
        If Len(num) = 8 Then
            wsh.SendKeys num
        Else
            wsh.SendKeys num & "{TAB}"
        End If
        WScript.Sleep 60
        batch = batch + 1
    Loop

    SK "~", 1200   ' Enter + wait for screen to process

    offset = offset + batch

    If offset < total Then
        ' Overflow: re-navigate for next screenful
        SK "{F5}",       600
        SK "7~",         600
        SK "Y",          200
        SK "N",          200
        SK "Y",          200
        SK "~",          600
        SK "{TAB}",      200
        SK "S",          200
        SK "{F5}",       600
        SK "{TAB}{TAB}", 200
    End If
Loop

' ─────────────────────────────────────────────────────────────────
' 9. MARK EMAIL AS READ AND FINISH
' ─────────────────────────────────────────────────────────────────
batchEmail.UnRead = False

MsgBox "All " & total & " number(s) submitted to DRIVES successfully.", _
       vbInformation, "DRIVES Batch Entry Complete"
