Attribute VB_Name = "modErrHandler"
Option Explicit

'The purpose of this module is as follows:
' * Provide at least some level of error tracking, rather than every error appearing as though it occured in the parent function
' * Allow functions to clean up their work when an error occurs, while still preserving the source of the original error

'Use the ErrorHandlerTemplates addin to ensure consistency of error-handling across the entire project.

Private Type Error
    Source As String
    Number As Long
    Description As String
End Type
Private ErrorLog() As Error
Private ErrorCount As Long

Sub ShowInfoMsg(e As String)
MsgBox e, vbInformation
End Sub

Sub ShowErrorMsg(e As String)
MsgBox e, vbCritical
End Sub

Function UNHANDLEDERROR(module As String, section As String, errobj As ErrObject, Optional INCLEANUP As Boolean = False) As Boolean
ReDim Preserve ErrorLog(ErrorCount)
With ErrorLog(ErrorCount)
    .Source = module & "." & section
    .Number = errobj.Number
    .Description = errobj.Description
    If INCLEANUP Then .Description = .Description & vbCrLf & "(Occurred during clean-up section)"
End With
ErrorCount = ErrorCount + 1
If INCLEANUP Then
    Err.Raise -1
Else
    Dim e As String, a As Long
    Const h As String = "Woops! An unhandled error has occurred. Under normal circumstances, this should never happen. Use Alt+PrtScrn to copy this, to paste into an email for the software developer."
    For a = 0 To ErrorCount - 1
        With ErrorLog(a)
            Select Case .Number
            Case -1
                e = e & vbCrLf & vbCrLf & "Error was ultimately caught by " & .Source & " error handler."
            Case 0
                e = e & vbCrLf & vbCrLf & .Source & ": Unknown error"
            Case 1
                e = e & vbCrLf & vbCrLf & .Source & ": " & .Description
            Case Else
                e = e & vbCrLf & vbCrLf & .Source & ": Error #" & .Number & ", " & .Description
            End Select
        End With
    Next a
    ErrorCount = 0
    MsgBox h & e, vbCritical
End If
UNHANDLEDERROR = True
End Function

Sub CAUSEERROR()
Err.Raise 1, , "Custom error message from CAUSEERROR() for debug purposes."
End Sub
