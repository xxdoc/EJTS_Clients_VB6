Attribute VB_Name = "modErrHandler"
Option Explicit

Public Const MAXERRS = 4    'Max number of errors allowed in one procedure before the error is bubbled up to the parent procedure

Sub ShowInfoMsg(e$)
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
MsgBox e$, vbInformation
End Sub

Sub ShowErrorMsg(e$)
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
MsgBox e$, vbCritical
End Sub

Sub UNHANDLEDERROR(section$)
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
Dim e$
Const h = "Woops! An unhandled error has occurred. Under normal circumstances, this should never happen. Please contact the software developer and report the error details below..." & vbCrLf & vbCrLf
If Err.Number = 0 Then
    e$ = "Error #0 (no details provided)"
ElseIf Err.Number = 1 Then
    e$ = Err.Description
Else
    e$ = "Error #" & Err.Number & ", " & Err.Description
End If

MsgBox h & "PROCEDURE: " & section$ & vbCrLf & _
           "DESCRIPTION: " & e$, vbCritical
End Sub

Sub CAUSEERROR()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
Err.Raise 1, , "Custom error message from CAUSEERROR() for debug purposes."
End Sub

