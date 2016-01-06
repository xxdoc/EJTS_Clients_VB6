Attribute VB_Name = "modMain"
Option Explicit
Private Const MOD_NAME = "modMain"

'MANIFEST HANDLER CODE FROM THE MANIFEST CREATOR ADD-IN>
Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
'<MANIFEST HANDLER CODE FROM THE MANIFEST CREATOR ADD-IN

'These must be public in order for the modules to access them>
Public DEBUGMODE As Boolean
Public FlashStopTime As Single
Public Const FlashDuration = 5     'Seconds
Public Enum enumScheduleMode
    sView
    sCreate
    sReschedule
End Enum

Public GS As CGlobalSettings
Public GlobalSettings() As Setting
Public GlobalSettings_Count As Long
Public GSChanged As Boolean
Public GSLoaded As Boolean

Public AppPath As String
Public DataFilesPath As String
Public FileToOpen_Year&
Public FileToOpen_OpenReadOnly As Boolean

Public ApptBeingRescheduled As Appointment
'<These must be public in order for the modules to access them

'EHT=Standard
'MANIFEST HANDLER CODE FROM THE MANIFEST CREATOR ADD-IN>
Sub Main()
On Error GoTo ERR_HANDLER

Dim iccex As InitCommonControlsExStruct, hMod As Long
' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
Const ICC_ANIMATE_CLASS As Long = &H80&
Const ICC_BAR_CLASSES As Long = &H4&
Const ICC_COOL_CLASSES As Long = &H400&
Const ICC_DATE_CLASSES As Long = &H100&
Const ICC_HOTKEY_CLASS As Long = &H40&
Const ICC_INTERNET_CLASSES As Long = &H800&
Const ICC_LINK_CLASS As Long = &H8000&
Const ICC_LISTVIEW_CLASSES As Long = &H1&
Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
Const ICC_PROGRESS_CLASS As Long = &H20&
Const ICC_TAB_CLASSES As Long = &H8&
Const ICC_TREEVIEW_CLASSES As Long = &H2&
Const ICC_UPDOWN_CLASS As Long = &H10&
Const ICC_USEREX_CLASSES As Long = &H200&
Const ICC_STANDARD_CLASSES As Long = &H4000&
Const ICC_WIN95_CLASSES As Long = &HFF&
Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

With iccex
   .lngSize = LenB(iccex)
   .lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
   ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
   ' example if using CommonControls v5.0 Progress bar:
   ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
End With
On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
InitCommonControlsEx iccex
If Err Then
    InitCommonControls ' try Win9x version
    Err.Clear
End If
On Error GoTo ERR_HANDLER
'... show your main form next (i.e., Form1.Show)
Main_AfterManifestHandling
If hMod Then FreeLibrary hMod

'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Main", Err
End Sub
'<MANIFEST HANDLER CODE FROM THE MANIFEST CREATOR ADD-IN

'EHT=Standard
Sub Main_AfterManifestHandling()
On Error GoTo ERR_HANDLER

Dim c$(), a

If RunningFromIDE Then DEBUGMODE = True
'DEBUGMODE = True        '[Mark]

'Location of exe
AppPath = AddTrailingSlash(App.Path)

'Check if SubClassingForVB.dll is registered
TryAgain:
If Not Is_SubClassingForVBdll_Registered Then
    If MsgBox("SubClassingForVB.dll has not been registered yet. Would you like to register it now?", vbQuestion Or vbOKCancel) = vbOK Then
        Shell "regsvr32 /s SubClassingForVB.dll", vbHide
        Sleep 2000
        GoTo TryAgain
    Else
        End
    End If
End If

If App.PrevInstance Then
    If MsgBox("Another instance is already running. Continue anyway?", vbCritical Or vbOKCancel) = vbCancel Then
        End
    End If
End If

'This must be after AppPath is set
Set GS = New CGlobalSettings        'This will call Main_Unload later
LoadGlobalSettings

'Location of data files, snapshots
DataFilesPath = DB_GetSetting(ActiveDBInstance, "GLOBAL_DataFolder")
If Not FolderExists(DataFilesPath) Then DataFilesPath = AppPath & "Data Files\"
If Not FolderExists(DataFilesPath) Then DataFilesPath = AppPath

'Process command line
c$ = Split(Command$, " ")
For a = 0 To UBound(c$)
    If LCase$(c$(a)) = "readonly" Then
        FileToOpen_OpenReadOnly = True
    ElseIf LCase$(c$(a)) = "auto" Then
        FileToOpen_Year = Year(Date) - 1
    ElseIf IsNumeric(c$(a)) Then
        FileToOpen_Year = Val(c$(a))
    End If
Next a

Set frmStart = New frmStart
frmStart.Form_Show      'Code will continue after this (not modal)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Main_AfterManifestHandling", Err
End Sub

'EHT=Standard
Sub Main_Unload()
On Error GoTo ERR_HANDLER

SaveGlobalSettings

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Main_Unload", Err
End Sub

'EHT=Custom
Function Is_SubClassingForVBdll_Registered() As Boolean
On Error GoTo e
Dim c As SubClass
Set c = New SubClass
Is_SubClassingForVBdll_Registered = True
Exit Function
e:
End Function
