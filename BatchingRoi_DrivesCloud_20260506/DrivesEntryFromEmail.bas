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
