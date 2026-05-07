Option Explicit

' TEST_MODE = True  → numbers passed to HTML page via URL query string.
'                     Edge auto-fills the form. User presses ENTER once to submit.
'                     No keystrokes sent from VBA at all.
' TEST_MODE = False → keybd_event hardware keystrokes sent to DRIVES Java terminal.
'                     (Java terminals accept OS-level hardware injection unlike Edge.)

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep       Lib "kernel32" (ByVal ms As Long)
    Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal vk As Byte, ByVal scan As Byte, ByVal flags As Long, ByVal extra As Long)
#Else
    Private Declare Sub Sleep       Lib "kernel32" (ByVal ms As Long)
    Private Declare Sub keybd_event Lib "user32" _
        (ByVal vk As Byte, ByVal scan As Byte, ByVal flags As Long, ByVal extra As Long)
#End If

Private Const TEST_MODE    As Boolean = True
Private Const TEST_HTML    As String  = "file:///C:/Users/zmasters1/Downloads/batch-driving-record.html"
Private Const DRIVES_TITLE As String  = "DRIVES"

Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_RETURN       As Byte = &HD
Private Const VK_TAB          As Byte = &H9
Private Const VK_F5           As Byte = &H74

Private m_wsh    As Object
Private m_target As String

' ── DRIVES-mode helpers (hardware keystrokes, ignored in test mode) ──
Private Sub PressVK(vk As Byte)
    keybd_event vk, 0, 0, 0
    Sleep 20
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
    Sleep 20
End Sub

Private Sub TypeText(s As String)
    Dim i As Long, c As String, vk As Byte, valid As Boolean
    For i = 1 To Len(s)
        c = UCase(Mid(s, i, 1))
        valid = False
        If c >= "0" And c <= "9" Then vk = CByte(Asc(c)) : valid = True
        If c >= "A" And c <= "Z" Then vk = CByte(Asc(c)) : valid = True
        If valid Then PressVK vk
        Sleep 30
    Next i
End Sub

Private Sub NavTo(text As String, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    TypeText text
    Sleep 30
    PressVK VK_RETURN
    Sleep waitMs
End Sub

Private Sub TextTo(text As String, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    TypeText text
    Sleep waitMs
End Sub

Private Sub KeyTo(vk As Byte, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    PressVK vk
    Sleep waitMs
End Sub

' ── MAIN ENTRY POINT  —  Alt+F8 → DrivesEntry → Run ─────────────
Public Sub DrivesEntry()

    ' ── 1. FIND BATCH EMAIL ──────────────────────────────────────
    Dim olApp As Object, ns As Object, inbox As Object
    Dim msg As Object, batchEmail As Object

    Set olApp = CreateObject("Outlook.Application")
    Set ns    = olApp.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(6)  ' olFolderInbox
    inbox.Items.Sort "[ReceivedTime]", True

    Set batchEmail = Nothing
    For Each msg In inbox.Items
        If msg.UnRead And InStr(msg.Subject, "DRIVES Batching -") > 0 Then
            Set batchEmail = msg
            Exit For
        End If
    Next

    If batchEmail Is Nothing Then
        MsgBox "No unread batch email found. Run the cloud flow first.", _
               vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If

    ' ── 2. EXTRACT NUMBERS FROM <pre>...</pre> ───────────────────
    Dim re As Object, mts As Object, preRaw As String
    Set re        = CreateObject("VBScript.RegExp")
    re.Pattern    = "<pre[^>]*>([\s\S]*?)</pre>"
    re.IgnoreCase = True
    Set mts = re.Execute(batchEmail.HTMLBody)

    If mts.Count = 0 Then
        MsgBox "No number list found in the email body.", vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If

    preRaw = mts(0).SubMatches(0)
    re.Pattern = "<[^>]+>" : re.Global = True
    preRaw = re.Replace(preRaw, "")
    preRaw = Replace(Replace(Replace(Replace(preRaw, _
             "&amp;", "&"), "&lt;", "<"), "&gt;", ">"), "&nbsp;", " ")
    preRaw = Replace(Replace(preRaw, vbCrLf, vbLf), vbCr, vbLf)

    Dim rawLines() As String : rawLines = Split(preRaw, vbLf)
    Dim nums()     As String : ReDim nums(UBound(rawLines))
    Dim n As Long, i As Long, ln As String

    For i = 0 To UBound(rawLines)
        ln = Trim(rawLines(i))
        If ln <> "" Then
            nums(n) = ln
            n = n + 1
        End If
    Next i

    If n = 0 Then
        MsgBox "No numbers found in the email.", vbExclamation, "DRIVES Batch Entry"
        Exit Sub
    End If
    ReDim Preserve nums(n - 1)

    ' ── 3. CONFIRM ───────────────────────────────────────────────
    Dim confirmMsg As String
    If TEST_MODE Then
        confirmMsg = n & " number(s) found." & vbCrLf & vbCrLf & _
                     "Click OK to open the test terminal in Edge." & vbCrLf & _
                     "The form will be pre-filled — press ENTER to submit."
    Else
        confirmMsg = n & " number(s) found." & vbCrLf & vbCrLf & _
                     "Make sure DRIVES is open on the Main Menu." & vbCrLf & _
                     "Click OK to begin entering numbers into DRIVES."
    End If

    If MsgBox(confirmMsg, vbOKCancel + vbInformation, "DRIVES Batch Entry") = vbCancel Then
        Exit Sub
    End If

    ' ── 4A. TEST MODE: open HTML with numbers pre-filled via URL ─
    If TEST_MODE Then

        Dim numsStr As String
        For i = 0 To n - 1
            If i > 0 Then numsStr = numsStr & ","
            numsStr = numsStr & nums(i)
        Next i

        Dim url As String
        url = TEST_HTML & "?nums=" & numsStr

        Shell "cmd /c start """" """ & url & """", vbHide

        batchEmail.UnRead = False

        MsgBox n & " number(s) loaded into the form in Edge." & vbCrLf & _
               "The form is pre-filled — press ENTER to submit.", _
               vbInformation, "DRIVES Batch Entry"

    ' ── 4B. DRIVES MODE: hardware keystrokes into Java terminal ──
    Else

        Set m_wsh  = CreateObject("WScript.Shell")
        m_target   = DRIVES_TITLE

        Application.WindowState = xlMinimized
        Sleep 300

        Dim focused As Boolean : focused = False
        Dim attempt As Long
        For attempt = 1 To 20
            focused = m_wsh.AppActivate(m_target)
            If focused Then Exit For
            Sleep 500
        Next attempt

        If Not focused Then
            Application.WindowState = xlNormal
            MsgBox "Could not find DRIVES window: """ & DRIVES_TITLE & """." & vbCrLf & _
                   "Make sure DRIVES is open.", vbExclamation, "DRIVES Batch Entry"
            Exit Sub
        End If

        Sleep 600

        ' Navigate to Batch Driving Record Request
        NavTo "3", 400 : NavTo "5", 400 : NavTo "14", 400 : NavTo "7", 400

        ' Header: Y / N / Y / Enter
        TextTo "Y", 200 : TextTo "N", 200 : TextTo "Y", 200
        KeyTo VK_RETURN, 500

        ' Email select: Tab → S → F5 → Tab → Tab
        KeyTo VK_TAB, 200 : TextTo "S", 200
        KeyTo VK_F5, 500 : KeyTo VK_TAB, 100 : KeyTo VK_TAB, 200

        ' Enter numbers into grid
        Dim maxBatch As Long : maxBatch = 90
        Dim offset   As Long : offset   = 0
        Dim batch    As Long
        Dim num      As String

        Do While offset < n
            batch = 0
            Do While batch < maxBatch And (offset + batch) < n
                num = nums(offset + batch)
                m_wsh.AppActivate m_target : Sleep 40
                TypeText num
                If Len(num) < 8 Then PressVK VK_TAB
                Sleep 50
                batch = batch + 1
            Loop

            KeyTo VK_RETURN, 1200
            offset = offset + batch

            If offset < n Then
                KeyTo VK_F5,    600 : NavTo "7",  600
                TextTo "Y", 200 : TextTo "N", 200 : TextTo "Y", 200
                KeyTo VK_RETURN, 600
                KeyTo VK_TAB, 200 : TextTo "S", 200
                KeyTo VK_F5, 600 : KeyTo VK_TAB, 100 : KeyTo VK_TAB, 200
            End If
        Loop

        batchEmail.UnRead = False
        Application.WindowState = xlNormal
        MsgBox "All " & n & " number(s) submitted to DRIVES successfully.", _
               vbInformation, "DRIVES Batch Entry"

    End If

End Sub

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare PtrSafe Sub mouse_event Lib "user32" _
        (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Sub mouse_event Lib "user32" _
        (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const KEYEVENTF_KEYUP      As Long = &H2
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP   As Long = &H4
Private Const SM_CXSCREEN          As Long = 0
Private Const SM_CYSCREEN          As Long = 1

Private Const VK_RETURN As Byte = &HD
Private Const VK_TAB    As Byte = &H9
Private Const VK_F5     As Byte = &H74

' ─────────────────────────────────────────────────────────────────
' TEST MODE
'   True  → opens batch-driving-record.html and targets "Test Terminal"
'   False → targets real DRIVES Java terminal
' ─────────────────────────────────────────────────────────────────
Private Const TEST_MODE    As Boolean = True
Private Const TEST_HTML    As String  = "file:///C:/Users/zmasters1/Downloads/batch-driving-record.html"
Private Const DRIVES_TITLE As String  = "DRIVES"
Private Const TEST_TITLE   As String  = "Test Terminal"

' Module-level so helpers can see them
Private m_wsh    As Object
Private m_target As String

' Click the centre of the screen — lands on the terminal, triggering our
' mousedown handler which explicitly calls input.focus() on the active input.
Private Sub ClickCenter()
    Dim cx As Long, cy As Long
    cx = GetSystemMetrics(SM_CXSCREEN) \ 2
    cy = GetSystemMetrics(SM_CYSCREEN) \ 2
    SetCursorPos cx, cy
    Sleep 80
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Sleep 50
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Sleep 200
End Sub

' Press and release one virtual key at hardware level
Private Sub PressVK(vk As Byte)
    keybd_event vk, 0, 0, 0
    Sleep 20
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
    Sleep 20
End Sub

' Type each alphanumeric character as hardware keystrokes
Private Sub TypeText(s As String)
    Dim i As Long, c As String, vk As Byte, ok As Boolean
    For i = 1 To Len(s)
        c = UCase(Mid(s, i, 1))
        ok = True
        Select Case c
            Case "0" To "9" : vk = CByte(Asc(c))   ' 0x30-0x39
            Case "A" To "Z" : vk = CByte(Asc(c))   ' 0x41-0x5A
            Case Else        : ok = False
        End Select
        If ok Then PressVK vk
        Sleep 30
    Next i
End Sub

' Activate window, type text only, then wait
Private Sub TextTo(text As String, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    TypeText text
    Sleep waitMs
End Sub

' Activate window, type text then Enter, then wait
Private Sub NavTo(text As String, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    TypeText text : Sleep 30
    PressVK VK_RETURN
    Sleep waitMs
End Sub

' Activate window, press a single special key, then wait
Private Sub KeyTo(vk As Byte, waitMs As Long)
    m_wsh.AppActivate m_target : Sleep 80
    PressVK vk
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

    Sleep 600   ' let Edge settle

    ' Click centre of screen — our HTML mousedown handler calls input.focus()
    ' which is the only reliable way to get DOM focus into Edge from VBA.
    ClickCenter
    Sleep 300

    ' ── 5. NAVIGATE TO BATCH DRIVING RECORD REQUEST SCREEN ───────
    NavTo "3",  400
    NavTo "5",  400
    NavTo "14", 400
    NavTo "7",  400

    ' ── 6. HEADER: Y → N → Y → Enter ────────────────────────────
    TextTo "Y", 200
    TextTo "N", 200
    TextTo "Y", 200
    KeyTo VK_RETURN, 500

    ' ── 7. EMAIL SELECT: Tab → S → F5 → Tab → Tab ───────────────
    KeyTo VK_TAB,    200
    TextTo "S",      200
    KeyTo VK_F5,     500
    KeyTo VK_TAB,    100
    KeyTo VK_TAB,    200

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
            m_wsh.AppActivate m_target : Sleep 50
            TypeText num
            If Len(num) < 8 Then PressVK VK_TAB
            Sleep 60
            batch = batch + 1
        Loop

        KeyTo VK_RETURN, 1200   ' submit screenful

        offset = offset + batch

        If offset < total Then
            ' Overflow — re-navigate for next screenful
            KeyTo VK_F5,     600
            NavTo "7",       600
            TextTo "Y",      200
            TextTo "N",      200
            TextTo "Y",      200
            KeyTo VK_RETURN, 600
            KeyTo VK_TAB,    200
            TextTo "S",      200
            KeyTo VK_F5,     600
            KeyTo VK_TAB,    100
            KeyTo VK_TAB,    200
        End If

    Loop

    ' ── 9. MARK EMAIL AS READ AND FINISH ─────────────────────────
    batchEmail.UnRead = False

    Application.WindowState = xlNormal
    MsgBox "All " & total & " number(s) submitted to DRIVES successfully.", _
           vbInformation, "DRIVES Batch Entry Complete"

End Sub
