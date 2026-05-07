Option Explicit

' Millisecond-precision sleep (works on all modern Office installs)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ─────────────────────────────────────────────────────────────────
' TEST MODE
'   True  → opens batch-driving-record.html and targets "Test Terminal"
'   False → targets real DRIVES Java terminal
' ─────────────────────────────────────────────────────────────────
Private Const TEST_MODE    As Boolean = True
Private Const TEST_HTML    As String  = "file:///C:/Users/zmasters1/Downloads/batch-driving-record.html"
Private Const DRIVES_TITLE As String  = "DRIVES"
Private Const TEST_TITLE   As String  = "Test Terminal"

' Module-level so SK helper can see them
Private m_wsh    As Object
Private m_target As String

Private Sub SK(keys As String, waitMs As Long)
    m_wsh.AppActivate m_target
    m_wsh.SendKeys keys
    Sleep waitMs
End Sub

' ─────────────────────────────────────────────────────────────────
' MAIN ENTRY POINT  —  Tools > Macros > DrivesEntry > Run
' ─────────────────────────────────────────────────────────────────
Public Sub DrivesEntry()

    If TEST_MODE Then m_target = TEST_TITLE Else m_target = DRIVES_TITLE

    ' ── 1. FIND THE BATCH EMAIL ──────────────────────────────────
    Dim outlook As Object, ns As Object, inbox As Object
    Dim msg As Object, batchEmail As Object

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
               "Run the cloud flow first.", vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If

    ' ── 2. EXTRACT NUMBERS FROM <pre>...</pre> ───────────────────
    Dim htmlBody As String, preRaw As String
    htmlBody = batchEmail.HTMLBody

    Dim re As Object
    Set re        = CreateObject("VBScript.RegExp")
    re.Pattern    = "<pre[^>]*>([\s\S]*?)</pre>"
    re.IgnoreCase = True

    Dim mts As Object
    Set mts = re.Execute(htmlBody)

    If mts.Count = 0 Then
        MsgBox "No number list found in the email body.", _
               vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If

    preRaw = mts(0).SubMatches(0)

    ' Strip HTML tags
    re.Pattern = "<[^>]+>"
    re.Global  = True
    preRaw = re.Replace(preRaw, "")

    ' Decode common entities
    preRaw = Replace(preRaw, "&amp;",  "&")
    preRaw = Replace(preRaw, "&lt;",   "<")
    preRaw = Replace(preRaw, "&gt;",   ">")
    preRaw = Replace(preRaw, "&nbsp;", " ")

    ' Normalise line endings
    preRaw = Replace(preRaw, vbCrLf, vbLf)
    preRaw = Replace(preRaw, vbCr,   vbLf)

    Dim rawLines() As String
    rawLines = Split(preRaw, vbLf)

    Dim numbers() As String
    Dim numCount As Long : numCount = 0
    ReDim numbers(UBound(rawLines))

    Dim ln As String
    Dim i As Long
    For i = 0 To UBound(rawLines)
        ln = Trim(rawLines(i))
        If ln <> "" Then
            numbers(numCount) = ln
            numCount = numCount + 1
        End If
    Next i

    If numCount = 0 Then
        MsgBox "The batch email was found but contained no numbers.", _
               vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If
    ReDim Preserve numbers(numCount - 1)

    ' ── 3. CONFIRM ───────────────────────────────────────────────
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox( _
        "Ready to enter " & numCount & " number(s) into DRIVES." & vbCrLf & vbCrLf & _
        "Make sure:" & vbCrLf & _
        "  1. DRIVES Java terminal is open" & vbCrLf & _
        "  2. You are on the Main Menu (screen 0)" & vbCrLf & vbCrLf & _
        "Click OK to begin.", _
        vbOKCancel + vbInformation, "DRIVES Batch Entry")

    If confirm = vbCancel Then Exit Sub

    ' ── 4. FOCUS TARGET WINDOW ───────────────────────────────────
    Set m_wsh = CreateObject("WScript.Shell")

    ' Minimise Excel so Windows focus-stealing prevention doesn't block AppActivate
    Application.WindowState = xlMinimized
    Sleep 300

    ' Try to focus the window — if not found and in test mode, launch the HTML page
    Dim focused As Boolean : focused = False
    Dim attempt As Long

    focused = m_wsh.AppActivate(m_target)

    If Not focused And TEST_MODE Then
        ' Window not open yet — launch it then wait for it to load
        Shell "cmd /c start """" """ & TEST_HTML & """", vbHide
        For attempt = 1 To 20
            Sleep 500
            focused = m_wsh.AppActivate(m_target)
            If focused Then Exit For
        Next attempt
    ElseIf Not focused Then
        ' DRIVES not found — retry a few times in case it's slow to respond
        For attempt = 1 To 10
            Sleep 500
            focused = m_wsh.AppActivate(m_target)
            If focused Then Exit For
        Next attempt
    End If

    If Not focused Then
        Application.WindowState = xlNormal
        MsgBox "Could not find window: """ & m_target & """." & vbCrLf & _
               "Make sure the window is open and the title matches.", _
               vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If

    Sleep 1000  ' wait for Edge to restore DOM focus to the page input

    ' ── 5. NAVIGATE TO BATCH DRIVING RECORD REQUEST SCREEN ───────
    SK "3~",  400
    SK "5~",  400
    SK "14~", 400
    SK "7~",  400

    ' ── 6. HEADER: Y → N → Y → Enter ────────────────────────────
    SK "Y", 200
    SK "N", 200
    SK "Y", 200
    SK "~", 500

    ' ── 7. EMAIL SELECT: Tab → S → F5 → Tab → Tab ───────────────
    SK "{TAB}",      200
    SK "S",          200
    SK "{F5}",       500
    SK "{TAB}{TAB}", 200

    ' ── 8. TYPE NUMBERS INTO GRID ────────────────────────────────
    Dim maxPerScreen As Long : maxPerScreen = 90
    Dim offset As Long       : offset       = 0
    Dim total As Long        : total        = numCount
    Dim batch As Long
    Dim num As String

    Do While offset < total

        batch = 0
        Do While batch < maxPerScreen And (offset + batch) < total
            num = numbers(offset + batch)
            If Len(num) = 8 Then
                m_wsh.SendKeys num          ' auto-advances
            Else
                m_wsh.SendKeys num & "{TAB}"  ' manual advance
            End If
            Sleep 60
            batch = batch + 1
        Loop

        SK "~", 1200   ' Enter + wait for screen to process

        offset = offset + batch

        If offset < total Then
            ' Overflow — re-navigate for next screenful
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

    ' ── 9. MARK EMAIL AS READ AND FINISH ─────────────────────────
    batchEmail.UnRead = False

    Application.WindowState = xlNormal
    MsgBox "All " & total & " number(s) submitted to DRIVES successfully.", _
           vbInformation, "DRIVES Batch Entry Complete"

End Sub
