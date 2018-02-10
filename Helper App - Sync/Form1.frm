VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sync"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   5910
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject, CL() As String, Cancel As Boolean, E As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Cancel = True Then Unload Me
    Cancel = True
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Show
CL() = Split(Command$, ">")
If FSO.FolderExists(CL(0)) = False Then MsgBox "Error finding source folder: " & CL(0): Unload Me: Exit Sub
If FSO.FolderExists(CL(1)) = False Then MsgBox "Error finding destination folder: " & CL(1): Unload Me: Exit Sub
Run FSO.GetFolder(CL(0)), FSO.GetFolder(CL(1))
If Cancel = True Then Text1.Text = Text1.Text & vbCrLf & "Canceled"
Text1.Text = Text1.Text & vbCrLf & "Done."
Text1.SelStart = Len(Text1.Text)
Cancel = True
End Sub

Sub Run(foS As Folder, foD As Folder)
On Error GoTo Error
DoEvents: If Cancel = True Then Exit Sub
Dim fi As File, sfo As Folder
Text1.Text = Text1.Text & vbCrLf
If foS.Files.Count <> 0 Then
    For Each fi In foS.Files
        If (fi.Attributes And Archive) = Archive Then
            DoEvents: If Cancel = True Then Exit Sub
            Text1.Text = Text1.Text & "Copying file: " & fi.Name & " . . . "
            Text1.SelStart = Len(Text1.Text)
            DoEvents
            E = False
            FSO.CopyFile fi.Path, foD & "\"
            If E = True Then GoTo SKIPA Else Text1.Text = Text1.Text & "          "
            
            DoEvents: If Cancel = True Then Exit Sub
            Text1.Text = Text1.Text & "Removing archive attribute . . . "
            Text1.SelStart = Len(Text1.Text)
            DoEvents
            E = False
            fi.Attributes = Not ((Not fi.Attributes) Or Archive)
            If E = False Then Text1.Text = Text1.Text & vbCrLf
SKIPA:
        End If
    Next fi
End If

If foS.SubFolders.Count <> 0 Then
    For Each sfo In foS.SubFolders
        If FSO.FolderExists(foD.Path & "\" & sfo.Name) = False Then
            DoEvents: If Cancel = True Then Exit Sub
            Text1.Text = Text1.Text & "Creating Folder: " & foD.Path & "\" & sfo.Name & " . . . "
            Text1.SelStart = Len(Text1.Text)
            DoEvents
            E = False
            FSO.CreateFolder (foD.Path & "\" & sfo.Name)
            If E = False Then Text1.Text = Text1.Text & vbCrLf: Text1.SelStart = Len(Text1.Text)
        End If
        Text1.Text = Text1.Text & "Opening Folder: " & foD.Path & "\" & sfo.Name & " . . . "
        Text1.SelStart = Len(Text1.Text)
        Run sfo, FSO.GetFolder(foD.Path & "\" & sfo.Name)
    Next sfo
End If
Exit Sub

Error:
Text1.Text = Text1.Text & "Failed" & vbCrLf
Text1.SelStart = Len(Text1.Text)
E = True
Resume Next
End Sub
