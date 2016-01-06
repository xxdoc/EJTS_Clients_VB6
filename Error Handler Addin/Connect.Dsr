VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   16050
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   23505
   _ExtentX        =   41460
   _ExtentY        =   28310
   _Version        =   393216
   Description     =   "Ensures all procedures use consistent error handler code"
   DisplayName     =   "Error Handler Templates"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private VBInstance As VBIDE.VBE
Private cmdbar As CommandBar
Private frm As frmAddIn
Private WithEvents ButtonEvents As CommandBarEvents          'command bar event handler
Attribute ButtonEvents.VB_VarHelpID = -1

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
On Error GoTo e

Dim btn As CommandBarButton

Set VBInstance = Application

If ConnectMode <> ext_cm_External Then
    Set cmdbar = VBInstance.CommandBars.Add("ErrorHandlerTemplates", msoBarTop, , True)
    cmdbar.RowIndex = 2
    
    Set btn = cmdbar.Controls.Add(msoControlButton)
    btn.FaceId = 1048
    btn.Caption = "Check EH"
    btn.Style = msoButtonIconAndCaption
    Set ButtonEvents = VBInstance.Events.CommandBarEvents(btn)
    
    #If False Then
        cmdbar.Position = msoBarFloating
        Dim a&
        For a = 0 To 500
            Set btn = cmdbar.Controls.Add(msoControlButton)
            btn.Caption = a
            btn.FaceId = a
            DoEvents
        Next a
        cmdbar.Height = 1200
    #End If
    
    cmdbar.Visible = True
End If

Exit Sub
e:
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
On Error Resume Next

cmdbar.Delete

Unload frm
Set frm = Nothing
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub ButtonEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next

If frm Is Nothing Then
    Set frm = New frmAddIn
    Set frm.CurProject = VBInstance.ActiveVBProject
End If
frm.Show
End Sub
