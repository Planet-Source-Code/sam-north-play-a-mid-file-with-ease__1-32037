Attribute VB_Name = "MIDMod"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength _
    As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal _
    fdwError As Long, ByVal lpszErrorText As String, ByVal cchErrorText As Long) As Long

Public Sub OpenMID(MIDFile As String, MIDAlias As String)
Dim ErrCode As Long  ' MCI error code
ErrCode = mciSendString("open " & MIDFile & " alias " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub PlayMID(MIDAlias As String)
Dim ErrCode As Long  ' MCI error code

ErrCode = mciSendString("play " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub StopMID(MIDAlias As String)
' The position within the file does not move back to the beginning.
Dim ErrCode As Long  ' MCI error code
    
ErrCode = mciSendString("stop " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub CloseMID(MIDAlias As String)
Dim ErrCode As Long  ' MCI error code
ErrCode = mciSendString("close " & MIDAlias, "", 0, 0)
End Sub

Private Sub DisplayError(ByVal ErrCode As Long)
'This subroutine displays a dialog box with the text of the MCI error.  There's
Dim errstr As String  ' MCI error message text
Dim retval As Long    ' return value

' Get a string explaining the MCI error.
errstr = Space(128)
retval = mciGetErrorString(ErrCode, errstr, Len(errstr))
' Remove the terminating null and empty space at the end.
errstr = Left(errstr, InStr(errstr, vbNullChar) - 1)

' Display error
retval = MsgBox(errstr, vbOKOnly Or vbCritical, "MCI")
End Sub
