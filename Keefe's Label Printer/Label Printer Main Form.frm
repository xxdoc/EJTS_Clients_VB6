VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label Printer"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "Label Printer Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtPaperSize 
      Height          =   285
      Left            =   5760
      TabIndex        =   14
      Text            =   "121"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   3765
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2550
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1335
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtPrintCount 
      Height          =   315
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "12"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbNames 
      Height          =   315
      ItemData        =   "Label Printer Main Form.frx":058A
      Left            =   120
      List            =   "Label Printer Main Form.frx":058C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
   Begin VB.Label Label2 
      Caption         =   "PaperSize:"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Count:"
      Height          =   195
      Left            =   5040
      TabIndex        =   13
      Top             =   765
      Width           =   630
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   510
      Width           =   4470
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   270
      TabIndex        =   11
      Top             =   735
      Width           =   4470
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   270
      TabIndex        =   8
      Top             =   960
      Width           =   4470
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   270
      TabIndex        =   10
      Top             =   1185
      Width           =   4470
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   270
      TabIndex        =   7
      Top             =   1410
      Width           =   4470
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   270
      TabIndex        =   9
      Top             =   1635
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4755
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tLabel
    LName As String
    LLabel(6) As String
End Type
Dim LData(100) As tLabel
Public LabelNumber As Integer, PrintCount As Integer

Private Sub cmbNames_Click()
LabelNumber = cmbNames.ItemData(cmbNames.ListIndex)
For a = 1 To 6
    lblLabel(a).Caption = LData(LabelNumber).LLabel(a)
Next a
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdPrint.Enabled = True
End Sub

Private Sub cmdPrint_Click()
m = frmMain.MousePointer
frmMain.MousePointer = 11

On Error Resume Next
Dim newval%, newvals$, newvalc As Currency
newval = txtPaperSize
Printer.PaperSize = newval
If Err.Number > 0 Then
    If MsgBox("Error setting Printer.PaperSize to " & newval & ". Continue anyway?", vbCritical Or vbYesNo Or vbDefaultButton1) = vbNo Then
        GoTo Cleanup
    End If
    Err.Clear
End If
newval = 1
Printer.ScaleMode = newval
If Err.Number > 0 Then
    If MsgBox("Error setting Printer.ScaleMode to " & newval & ". Continue anyway?", vbCritical Or vbYesNo Or vbDefaultButton1) = vbNo Then
        GoTo Cleanup
    End If
    Err.Clear
End If
newvals$ = "Sans Serif 10cpi"
Printer.Font.Name = newvals$
If Err.Number > 0 Then
    If MsgBox("Error setting Printer.Font.Name to '" & newvals$ & "'. Continue anyway?", vbCritical Or vbYesNo Or vbDefaultButton1) = vbNo Then
        GoTo Cleanup
    End If
    Err.Clear
End If
newvalc = 12
Printer.Font.Size = newvalc
If Err.Number > 0 Then
    If MsgBox("Error setting Printer.Font.Size to " & newvalc & ". Continue anyway?", vbCritical Or vbYesNo Or vbDefaultButton1) = vbNo Then
        GoTo Cleanup
    End If
    Err.Clear
End If
On Error GoTo e

For a = 1 To PrintCount
    For b = 1 To 6
        Printer.Print lblLabel(b).Caption
        Printer.CurrentY = Printer.CurrentY + 40
    Next b
Next a

Cleanup:
    Printer.EndDoc
    frmMain.MousePointer = m

Exit Sub
e:
    MsgBox "Error {" & Err.Description & "} occurred while printing."
    GoTo Cleanup
End Sub

Private Sub cmdOrder_Click()
X = cmbNames.ListIndex
If X <> -1 Then X = cmbNames.ItemData(X)
frmOrder.lstNames.Clear
For a = 0 To cmbNames.ListCount - 1
    frmOrder.lstNames.List(a) = cmbNames.List(a)
    frmOrder.lstNames.ItemData(a) = cmbNames.ItemData(a)
Next a
frmOrder.lstNames.ListIndex = cmbNames.ListIndex
frmOrder.Canceled = True
frmOrder.Show 1
If frmOrder.Canceled = True Then Exit Sub
For a = 0 To cmbNames.ListCount - 1
    cmbNames.List(a) = frmOrder.lstNames.List(a)
    cmbNames.ItemData(a) = frmOrder.lstNames.ItemData(a)
    If X <> -1 And cmbNames.ItemData(a) = X Then S = a
Next a
If X <> -1 Then cmbNames.ListIndex = S
End Sub

Private Sub cmdScan_Click()
m = frmMain.MousePointer
frmMain.MousePointer = 11
frmMain.Enabled = False
'search for papersize of 8.5x12
On Error Resume Next
For p = 0 To 512
    txtPaperSize = p: DoEvents
    Printer.PaperSize = p
    Debug.Print Printer.PaperSize; Printer.Height / 1440; Printer.Width / 1440
    If Printer.Height = 12 * 1440 Then
        If Printer.Width = 8.5 * 1440 Then
            GoTo Found
        End If
    End If
Next p
txtPaperSize = ""
MsgBox "8.5x12 paper size not found."
Found:
frmMain.MousePointer = m
frmMain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveDataFile
End Sub

Private Sub txtPrintCount_GotFocus()
txtPrintCount.SelStart = 0
txtPrintCount.SelLength = Len(txtPrintCount.Text)
End Sub

Private Sub txtPrintCount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPrintCount_Validate False: cmdPrint.SetFocus
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtPrintCount_Validate(Cancel As Boolean)
X = Val(txtPrintCount.Text)
If X > 9996 Then
    PrintCount = 9996
ElseIf X < 1 Then
    PrintCount = 1
Else
    PrintCount = X
End If
If PrintCount > 12 Then PrintCount = ((PrintCount - 1) \ 12 + 1) * 12
txtPrintCount.Text = Trim$(Str$(PrintCount))
End Sub

Private Sub cmdAdd_Click()
frmEdit.txtName.Text = ""
For a = 1 To 6
    frmEdit.txtLabel(a).Text = ""
Next a
frmEdit.Canceled = True
frmEdit.Caption = "Add"
frmEdit.Show 1
If frmEdit.Canceled = False Then
    LabelNumber = cmbNames.ListCount
    LData(LabelNumber).LName = frmEdit.txtName.Text
    cmbNames.AddItem LData(LabelNumber).LName, LabelNumber
    cmbNames.ItemData(LabelNumber) = LabelNumber
    For a = 1 To 6
        LData(LabelNumber).LLabel(a) = frmEdit.txtLabel(a).Text
        lblLabel(a).Caption = LData(LabelNumber).LLabel(a)
    Next a
    cmbNames.ListIndex = LabelNumber
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
    cmdPrint.Enabled = True
    cmdOrder.Enabled = True
End If
Unload frmEdit
End Sub

Private Sub cmdDelete_Click()
LabelNumber = cmbNames.ListIndex
LData(cmbNames.ItemData(LabelNumber)) = LData(cmbNames.ListCount - 1)
For a = 0 To cmbNames.ListCount - 1
    If cmbNames.ItemData(a) = cmbNames.ListCount - 1 Then
        cmbNames.ItemData(a) = cmbNames.ItemData(LabelNumber)
        Exit For
    End If
Next a
cmbNames.RemoveItem LabelNumber
If cmbNames.ListCount = 0 Then
    For a = 1 To 6
        lblLabel(a).Caption = ""
    Next a
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdPrint.Enabled = False
    cmdOrder.Enabled = False
Else
    If LabelNumber > cmbNames.ListCount - 1 Then LabelNumber = cmbNames.ListCount - 1
    cmbNames.ListIndex = LabelNumber
End If
End Sub

Private Sub cmdEdit_Click()
LabelNumber = cmbNames.ListIndex
X = cmbNames.ItemData(LabelNumber)
frmEdit.txtName.Text = LData(X).LName
For a = 1 To 6
    frmEdit.txtLabel(a).Text = LData(X).LLabel(a)
Next a
frmEdit.Canceled = True
frmEdit.Caption = "Edit"
frmEdit.Show 1
If frmEdit.Canceled = False Then
    LData(X).LName = frmEdit.txtName.Text
    cmbNames.List(LabelNumber) = LData(X).LName
    For a = 1 To 6
        LData(X).LLabel(a) = frmEdit.txtLabel(a).Text
        lblLabel(a).Caption = LData(X).LLabel(a)
    Next a
End If
Unload frmEdit
End Sub

Private Sub Form_Load()
LoadDataFile
PrintCount = 12
cmdOrder.Enabled = (cmbNames.ListCount > 0)
If cmbNames.ListCount > 0 Then cmbNames.ListIndex = 0
End Sub

Sub LoadDataFile()
On Error GoTo FileError
Open "Label Printer.dat" For Input As 1
Input #1, PS$
txtPaperSize.Text = PS$
Input #1, N
For a = 0 To N - 1
    Input #1, LData(a).LName
    cmbNames.List(a) = LData(a).LName
    cmbNames.ItemData(a) = a
    For b = 1 To 6
        Input #1, LData(a).LLabel(b)
    Next b
Next a
Close
On Error GoTo 0
Exit Sub
FileError:
MsgBox "Error loading data file."
Close
On Error GoTo 0
End Sub

Sub SaveDataFile()
On Error GoTo FileError
Open "Label Printer.dat" For Output As 1
Write #1, txtPaperSize.Text
Write #1, cmbNames.ListCount
For a = 0 To cmbNames.ListCount - 1
    X = cmbNames.ItemData(a)
    Write #1, LData(X).LName
    For b = 1 To 6
        Write #1, LData(X).LLabel(b)
    Next b
Next a
Close
On Error GoTo 0
Exit Sub
FileError:
MsgBox "Error saving data file."
Close
On Error GoTo 0
End Sub
