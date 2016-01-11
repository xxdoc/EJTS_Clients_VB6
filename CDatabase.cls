VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CDatabase"

'Error descriptions
Private Const errDBFileMissing = "Database file cannot be found."
Private Const errAlreadyLoaded = "A database file is already loaded. Cannot open another until the first one is closed."
Private Const errNotLoaded = "Database not loaded."
Private Const errNotCurrentVersion = "Database was not saved using the latest specifications, so it cannot be opened."
Private Const errDamagedDB = "Database file has been damaged."

'Internal variables
Private Const mCurSpecVersion As Long = 1
Private mDBFile As CMNMOD_CFileHandler

'Database attributes
Private mLoaded As Boolean
Private mIsWriteable As Boolean
Private mIsDirty As Boolean
Private mDBFilename As String
Private mLogFilename As String

'Core data
Private mClients() As CClient
Private mClientCount As Long
Private mAppointments() As CAppointment
Private mAppointmentCount As Long
Private mApptCliLinks() As CApptCliLink
Private mApptCliLinkCount As Long
Private mTaxReturns() As CTaxReturn
Private mTaxReturnCount As Long

'Events
Public Event OnCreateInitialDB()
Public Event OnFinishedLoading()
Public Event OnChange()

'#################################################################################
'Load & save functions
'#################################################################################

'EHT=Standard
Public Function ConnectToDatabase(vDBFilename As String, Optional vReadOnly As Boolean, Optional vCreateIfNotExists As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim p As Long, a As Long

'Initialize
If mLoaded Then Err.Raise 1, , errAlreadyLoaded
mIsWriteable = Not vReadOnly
mDBFilename = vDBFilename

'Calculate logfile name
p = InStrRev(vDBFilename, ".")
If p > 0 Then
    a = InStrRev(vDBFilename, "\")
    If a < p Then
        vDBFilename = Left$(vDBFilename, p - 1)
    End If
End If
mLogFilename = vDBFilename & ".log"

'Open the file
If FileExists(mDBFilename) Then
    If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing
    Set mDBFile = OpenFile(mDBFilename, mBinary_Input)
    
    '-------------------------
    
    'Database header
    If mDBFile.ReadLong <> mCurSpecVersion Then Err.Raise 1, , errNotCurrentVersion
    
    'Clients
    mClientCount = mDBFile.ReadLong
    ReDim mClients(mClientCount - 1)
    For a = 0 To mClientCount - 1
        Set mClients(a) = New CClient: mClients(a).Unserialize Me, mDBFile
    Next a
    
    'Appointments
    mAppointmentCount = mDBFile.ReadLong
    ReDim mAppointments(mAppointmentCount - 1)
    For a = 0 To mAppointmentCount - 1
        Set mAppointments(a) = New CAppointment: mAppointments(a).Unserialize Me, mDBFile
    Next a
    
    'ApptCliLinks
    mApptCliLinkCount = mDBFile.ReadLong
    ReDim mApptCliLinks(mApptCliLinkCount - 1)
    For a = 0 To mApptCliLinkCount - 1
        Set mApptCliLinks(a) = New CApptCliLink: mApptCliLinks(a).Unserialize Me, mDBFile
    Next a
    
    'TaxReturns
    mTaxReturnCount = mDBFile.ReadLong
    ReDim mTaxReturns(mTaxReturnCount - 1)
    For a = 0 To mTaxReturnCount - 1
        Set mTaxReturns(a) = New CTaxReturn: mTaxReturns(a).Unserialize Me, mDBFile
    Next a
    
    'Database footer
    If mDBFile.ReadLong <> mCurSpecVersion Then Err.Raise 1, , errDamagedDB
    
    '-------------------------

    mIsDirty = False
Else
    If vCreateIfNotExists Then
        CreateFile mDBFilename
        mIsDirty = False
        RaiseEvent OnCreateInitialDB
    Else
        Err.Raise 1, , errDBFileMissing
    End If
End If

'Success
mLoaded = True: ConnectToDatabase = mLoaded
RaiseEvent OnFinishedLoading

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ConnectToDatabase", Err
End Function

'EHT=Cleanup1
Public Function SaveChanges() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim tempfile As String, a As Long

If Not mLoaded Then Err.Raise 1, , errNotLoaded

If mIsWriteable Then
    If mIsDirty Then
        'Open the file
        Dim mTempFile As CMNMOD_CFileHandler
        tempfile = mDBFilename & ".sav"
        Set mTempFile = OpenFile(tempfile, mBinary_Output)
        
        '-------------------------
        
        'Database header
        mTempFile.WriteLong mCurSpecVersion
        
        'Clients
        mTempFile.WriteLong mClientCount
        For a = 0 To mClientCount - 1
            mClients(a).Serialize mTempFile
        Next a
        
        'Appointments
        mTempFile.WriteLong mAppointmentCount
        For a = 0 To mAppointmentCount - 1
            mAppointments(a).Serialize mTempFile
        Next a
        
        'ApptCliLinks
        mTempFile.WriteLong mApptCliLinkCount
        For a = 0 To mApptCliLinkCount - 1
            mApptCliLinks(a).Serialize mTempFile
        Next a
        
        'TaxReturns
        mTempFile.WriteLong mTaxReturnCount
        For a = 0 To mTaxReturnCount - 1
            mTaxReturns(a).Serialize mTempFile
        Next a
        
        'Database footer
        mTempFile.WriteLong mCurSpecVersion
        
        '-------------------------
        
        'Close and move the file
        mTempFile.CloseFile: Set mTempFile = Nothing        'Close temp file so it can be moved
        If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing     'Release lock
        RenameFile tempfile, mDBFilename, True              'Overwrite original with temp file
        Set mDBFile = OpenFile(mDBFilename, mBinary_Input)  'Restablish lock on file
        mIsDirty = False
    End If
    
    'Return value of True means that the database file now reflects the data (regardless if we had to write any data to make that happen)
    SaveChanges = True
End If

CLEANUP: INCLEANUP = True
    If Not mTempFile Is Nothing Then mTempFile.CloseFile: Set mTempFile = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SaveChanges", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Standard
Public Function DisconnectFromDatabase() As Boolean
On Error GoTo ERR_HANDLER

If mLoaded Then
    If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing     'Release lock
    mLoaded = False
    mDBFilename = ""
    mLogFilename = ""
    Erase mClients
    mClientCount = 0
    Erase mAppointments
    mAppointmentCount = 0
    Erase mApptCliLinks
    mApptCliLinkCount = 0
    Erase mTaxReturns
    mTaxReturnCount = 0
End If

'True just means that the database is successfully disconnected (even if it wasn't even connected in the first place)
DisconnectFromDatabase = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DisconnectFromDatabase", Err
End Function

'EHT=None
Private Sub Class_Terminate()
DisconnectFromDatabase
End Sub









'#################################################################################
'Database attributes
'#################################################################################

'EHT=None
Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

'EHT=None
Public Sub SetDirty()
mIsDirty = True
RaiseEvent OnChange
End Sub

'EHT=None
Public Sub ClearDirty()
mIsDirty = False
End Sub













'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function ClientByID(vClientID As Long) As CClient
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a&
For a = mClientCount - 1 To 0 Step -1
    If mClients(a).ID = vClientID Then
        Set ClientByID = mClients(a)
        Exit For
    End If
Next a
End Function

'EHT=None
Public Function AppointmentByID(vAppointmentID As Long) As CAppointment
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a&
For a = mAppointmentCount - 1 To 0 Step -1
    If mAppointments(a).ID = vAppointmentID Then
        Set AppointmentByID = mAppointments(a)
        Exit For
    End If
Next a
End Function

'EHT=None
Public Function AppointmentsByClient(vClientID As Long) As CAppointment()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a&, b&, c&, l() As CAppointment
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .ClientID = vClientID Then
            ReDim Preserve l(c)
            Set l(c) = AppointmentByID(.AppointmentID)
            c = c + 1
        End If
    End With
Next a
AppointmentsByClient = l
End Function

'EHT=None
Public Function ClientsByAppointment(vAppointmentID As Long) As CClient()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a&, b&, c&, l() As CClient
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .AppointmentID = vAppointmentID Then
            ReDim Preserve l(c)
            Set l(c) = ClientByID(.ClientID)
            c = c + 1
        End If
    End With
Next a
ClientsByAppointment = l
End Function
