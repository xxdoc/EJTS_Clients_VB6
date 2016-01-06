VERSION 5.00
Begin VB.Form frmAddIn 
   Caption         =   "Error Handler Templates"
   ClientHeight    =   8295
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4800
      Width           =   4695
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton btnApply 
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Tag             =   "Apply All Pending Changes"
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.ListBox lstPreview 
      Height          =   300
      IntegralHeight  =   0   'False
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Template
    Trigger As String
    
    Header As String
    HeaderWithWildcard As String
    HeaderLineCount As Long
    
    PreFooterTrigger As String
    PreFooter As String
    PreFooterWithWildcard As String
    PreFooterLineCount As Long
    
    Footer As String
    FooterWithWildcard As String
    FooterLineCount As Long
End Type
Private Templates() As Template
Private Const Wildcard_ProcName = "<PROCNAME>"
Private Const Wildcard_SubOrFunc = "<SUBFUNCPROP>"

Private Const ErrorHandlingConst = "MOD_NAME"

Private AnythingToApply As Boolean

Public CurProject As VBProject





Private Sub Form_Load()
On Error GoTo ERR_HANDLER

btnApply.Enabled = False
btnApply.Caption = "Loading templates..."
Me.Show
InitializeTemplates

btnApply.Caption = "Analyzing code..."
AnythingToApply = False
If UpdateErrorHandlers(False) Then
    If AnythingToApply Then
        btnApply.Enabled = True
        btnApply.Caption = btnApply.Tag
    Else
        MsgBox "No changes to be made", vbInformation
        Unload Me
    End If
Else
    btnApply.Caption = "ERRORS!"
End If

Exit Sub
ERR_HANDLER:
    MsgBox "ERROR initializing addin: " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub btnApply_Click()
btnApply.Enabled = False
btnApply.Caption = "Applying changes..."

If UpdateErrorHandlers(True) Then
    MsgBox "Success!", vbInformation
    Unload Me
Else
    btnApply.Caption = "ERRORS!"
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move (Me.ScaleWidth / 2) - (Frame1.Width / 2), Me.ScaleHeight - Frame1.Height - 8
lstPreview.Move 8, 8, Me.ScaleWidth - 16, Frame1.Top - 16
End Sub










Function CountLines(str As String) As Long
Dim a As Long
If Len(str) = 0 Then
    CountLines = 0
Else
    CountLines = 1
    Do
        a = InStr(a + 1, str, vbCrLf)
        If a = 0 Then Exit Do
        CountLines = CountLines + 1
    Loop
End If
End Function

Function ReplacePlaceholdersWithWildcards(str As String) As String
Const wc = "*"
ReplacePlaceholdersWithWildcards = Replace(Replace(str, Wildcard_ProcName, wc), Wildcard_SubOrFunc, wc)
End Function

Function ReplacePlaceholdersWithValues(str As String, procname As String, suborfunc As String) As String
ReplacePlaceholdersWithValues = Replace(Replace(str, Wildcard_ProcName, procname), Wildcard_SubOrFunc, suborfunc)
End Function

Function FindTextInCode(code As CodeModule, findstr As String, ByVal line As Long, ByVal linecount As Long) As Boolean
FindTextInCode = code.Find(findstr, line, 0, line + linecount, 0, False, True, False)
End Function

Function FindLineInCode(code As CodeModule, findstr As String, ByVal line As Long, ByVal linecount As Long) As Boolean
Dim col As Long, endline As Long
endline = line + linecount
Do
    'The +0 makes prevents the Find function from modifying our original 'endline' variable
    If code.Find(findstr, line, 0, endline + 0, 0, False, True, False) Then
        If code.Lines(line, 1) = findstr Then
            FindLineInCode = True
            Exit Do
        End If
    Else
        Exit Do
    End If
    line = line + 1
Loop
End Function

Function ErrorHandlingEnabled(code As CodeModule) As Boolean
On Error Resume Next
Dim m As Member
Set m = code.Members(ErrorHandlingConst)
ErrorHandlingEnabled = (m.Type = vbext_mt_Const) And (m.Scope = vbext_Private)
End Function










Sub InitializeTemplates()
Dim a As Long

ReDim Templates(4)
With Templates(0)
    .Trigger = "EHT=Custom"
End With
With Templates(1)
    .Trigger = "EHT=ResumeNext"
    .Header = "On Error Resume Next" & _
              vbCrLf
End With
With Templates(2)
    .Trigger = "EHT=Standard"
    .Header = "On Error GoTo ERR_HANDLER" & _
              vbCrLf
    .Footer = vbCrLf & _
              "Exit <SUBFUNCPROP>" & vbCrLf & _
              "ERR_HANDLER: UNHANDLEDERROR MOD_NAME, ""<PROCNAME>"", Err"
End With
With Templates(3)
    .Trigger = "EHT=Cleanup1"
    .Header = "On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean" & _
              vbCrLf
    .PreFooterTrigger = "CLEANUP: INCLEANUP = True"
    .PreFooter = vbCrLf & _
                 .PreFooterTrigger & vbCrLf & _
                 "    '(clean-up code)"
    .Footer = vbCrLf & _
              "Exit <SUBFUNCPROP>" & vbCrLf & _
              "ERR_HANDLER: UNHANDLEDERROR MOD_NAME, ""<PROCNAME>"", Err, INCLEANUP: Resume CLEANUP"
End With
With Templates(4)
    .Trigger = "EHT=Cleanup2"
    .Header = "On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean" & _
              vbCrLf
    .PreFooterTrigger = "CLEANUP: INCLEANUP = True"
    .PreFooter = vbCrLf & _
                 .PreFooterTrigger & vbCrLf & _
                 "    '(clean-up code)"
    .Footer = vbCrLf & _
              "Exit <SUBFUNCPROP>" & vbCrLf & _
              "ERR_HANDLER: UNHANDLEDERROR MOD_NAME, ""<PROCNAME>"", Err, INCLEANUP: HASERROR = True: Resume CLEANUP"
End With

For a = 0 To UBound(Templates)
    With Templates(a)
        If Len(.Trigger) = 0 Then Err.Raise 1, , "Template #" & a & " is invalid."
        .HeaderWithWildcard = ReplacePlaceholdersWithWildcards(.Header)
        .HeaderLineCount = CountLines(.Header)
        .PreFooterWithWildcard = ReplacePlaceholdersWithWildcards(.PreFooter)
        .PreFooterLineCount = CountLines(.PreFooter)
        .FooterWithWildcard = ReplacePlaceholdersWithWildcards(.Footer)
        .FooterLineCount = CountLines(.Footer)
    End With
Next a
End Sub

Function UpdateErrorHandlers(MakeChanges As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim component As VBComponent, code As CodeModule, proc As Member, procname As String, proclocation As Long
Dim dl As Long, t As Long
Dim atleastoneerror As Boolean

lstPreview.Clear

'For each code module...
For Each component In CurProject.VBComponents
    Select Case component.Type
    Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl
        Set code = component.CodeModule
        dl = code.CountOfDeclarationLines
        'If lstPreview.ListCount > 0 Then lstPreview.AddItem ""
        'If error handling in this module is enabled...
        If ErrorHandlingEnabled(code) Then
            lstPreview.AddItem "    ANALYZING " & code.Name & "..."
            'For each procedure group in the code module...
            For Each proc In component.CodeModule.Members
                Select Case proc.Type
                Case vbext_mt_Property, vbext_mt_Event, vbext_mt_Method
                    proclocation = proc.CodeLocation
                    procname = proc.Name
                    If proclocation > dl Then
                        'Find the Get/Let/Set procedures in the group, or simply the standard procedure...
                        For t = 0 To 3
                            'For each one...
                            If Not ProcessProcedure(code, procname, t, MakeChanges) Then atleastoneerror = True
                        Next
                    End If
                End Select
            Next
        Else
            lstPreview.AddItem "    ANALYZING " & code.Name & "... skipping, since constant " & ErrorHandlingConst & " is not defined."
        End If
    End Select
Next

If atleastoneerror Then
    lstPreview.AddItem ""
    lstPreview.AddItem "You must resolve all errors listed above before changes can be applied."
Else
    UpdateErrorHandlers = True
End If

Exit Function
ERR_HANDLER:
    If lstPreview.ListCount > 0 Then lstPreview.AddItem ""
    lstPreview.AddItem "Changes were only partially applied, due to an error in the UpdateErrorHandlers function."
    lstPreview.AddItem "ERROR: " & Err.Description
    lstPreview.TopIndex = lstPreview.ListCount - 1
End Function

Function ProcessProcedure(code As CodeModule, procname As String, proctype As vbext_ProcKind, MakeChanges As Boolean) As Boolean
Dim procname2 As String, procstart As Long, proclinecount As Long
Dim bodyfirst As Long, bodylast As Long
Dim subfuncprop As String, lns() As String
Dim oldstr As String, newstr As String, usingprefooter As Boolean, fc As Long, a As Long, curtpl As Long
Dim logstr As String

On Error Resume Next
procstart = code.ProcStartLine(procname, proctype)
If procstart = 0 Then
    ProcessProcedure = True
    Exit Function
End If
On Error GoTo ERR_HANDLER

procname2 = procname & Choose(proctype + 1, "", "[Let]", "[Set]", "[Get]")

'TryAgain:

'Determine which template to use
proclinecount = code.ProcCountLines(procname, proctype)
curtpl = -1
For a = 0 To UBound(Templates)
    If FindTextInCode(code, Templates(a).Trigger, procstart, proclinecount) Then
        curtpl = a
        Exit For
    End If
Next a
If curtpl < 0 Then
    'All procedures must specify an EHT=, so error at this point if we couldn't find a valid one
    'bodyfirst = code.ProcBodyLine(procname, proctype)
    'code.InsertLines bodyfirst, "'EHT=Standard"
    'GoTo TryAgain
    Err.Raise 1, , "Could not find a valid EHT= definition"
End If

With Templates(curtpl)
    'Find the first line of the procedure's content
    bodyfirst = code.ProcBodyLine(procname, proctype)
    Do
        oldstr = code.Lines(bodyfirst, 1)
        bodyfirst = bodyfirst + 1
        If Right$(oldstr, 2) = " _" Then
            subfuncprop = subfuncprop & Left$(oldstr, Len(oldstr) - 2)
        Else
            subfuncprop = subfuncprop & oldstr
            Exit Do
        End If
    Loop
    
    'Is it a Sub, Function, or Property?
    subfuncprop = Trim$(Left$(subfuncprop, InStr(subfuncprop, "(") - 1))
    lns = Split(subfuncprop)
    subfuncprop = "__?__"
    For a = 0 To UBound(lns)
        Select Case LCase$(lns(a))
        Case "sub", "function", "property"
            subfuncprop = lns(a)
        End Select
    Next a
    
    If .HeaderLineCount > 0 Then
        'Insert/update header
        oldstr = code.Lines(bodyfirst, .HeaderLineCount)
        newstr = ReplacePlaceholdersWithValues(.Header, procname2, subfuncprop)
        If oldstr Like .HeaderWithWildcard Then
            If oldstr <> newstr Then
                If MakeChanges Then
                    code.DeleteLines bodyfirst, .HeaderLineCount
                    code.InsertLines bodyfirst, newstr
                End If
                AnythingToApply = True
                logstr = logstr & IIf(Len(logstr) > 0, "; ", "") & "update header"
            End If
        Else
            If MakeChanges Then
                code.InsertLines bodyfirst, newstr
            End If
            AnythingToApply = True
            logstr = logstr & IIf(Len(logstr) > 0, "; ", "") & "add header"
        End If
    End If
    
    'Find the last line of the procedure's content
    proclinecount = code.ProcCountLines(procname, proctype)
    bodylast = procstart + proclinecount - 1
    Do Until Left$(Trim$(code.Lines(bodylast, 1)), 4) = "End "
        bodylast = bodylast - 1
    Loop
    bodylast = bodylast - 1
    
    If .FooterLineCount > 0 Then
        'Insert/update footer
        If .PreFooterLineCount > 0 Then usingprefooter = Not FindLineInCode(code, .PreFooterTrigger, procstart, proclinecount)
        If usingprefooter Then
            fc = .FooterLineCount + .PreFooterLineCount
            newstr = ReplacePlaceholdersWithValues(.PreFooter & vbCrLf & .Footer, procname2, subfuncprop)
        Else
            fc = .FooterLineCount
            newstr = ReplacePlaceholdersWithValues(.Footer, procname2, subfuncprop)
        End If
        oldstr = code.Lines(bodylast - fc + 1, fc)
        If oldstr Like .FooterWithWildcard Then
            If oldstr <> newstr Then
                If MakeChanges Then
                    code.DeleteLines bodylast - fc + 1, fc
                    code.InsertLines bodylast - fc + 1, newstr
                End If
                AnythingToApply = True
                logstr = logstr & IIf(Len(logstr) > 0, "; ", "") & "update " & IIf(usingprefooter, "pre-footer & ", "") & "footer"
            End If
        Else
            If MakeChanges Then
                code.InsertLines bodylast + 1, newstr
                bodylast = bodylast + fc
            End If
            AnythingToApply = True
            logstr = logstr & IIf(Len(logstr) > 0, "; ", "") & "add " & IIf(usingprefooter, "pre-footer & ", "") & "footer"
        End If
    End If
End With

If Len(logstr) > 0 Then
    lstPreview.AddItem code.Name & "." & procname2 & " - " & logstr
    lstPreview.TopIndex = lstPreview.ListCount - 1
End If

ProcessProcedure = True

Exit Function
ERR_HANDLER:
    lstPreview.AddItem code.Name & "." & procname2 & " - ERROR: " & Err.Description
    lstPreview.TopIndex = lstPreview.ListCount - 1
End Function

