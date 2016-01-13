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
Private Const errDBFileMissing = "Database file '%1%' could not be found (or was an empty file)."
Private Const errAlreadyLoaded = "A database file is already loaded. Cannot open another until the first one is closed."
Private Const errNotLoaded = "Database not loaded."
Private Const errSchemaNotSupported = "The database schema version of the specified database (%1%) is outside of the range supported by this program (%2% to %3%)."
Private Const errDamagedDB = "Database file '%1%' has been damaged."

'Internal variables
Private Const LowestSupportedSchemaVersion As Long = 1
Private Const LatestSchemaVersion As Long = 1
Private mDBFile As CMNMOD_CFileHandler

'Database attributes (accessible through class properties)
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

Dim p As Long, a As Long, schemaversion As Long

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

If FileExists(mDBFilename) Then
    'Open the file
    Set mDBFile = OpenFile(mDBFilename, mBinary_Input)
    
    '-------------------------
    
    'Database header
    schemaversion = mDBFile.ReadLong
    If (schemaversion < LowestSupportedSchemaVersion) Or (schemaversion > LatestSchemaVersion) Then Err.Raise 1, , FormatErrorMsg(errSchemaNotSupported, schemaversion, LowestSupportedSchemaVersion, LatestSchemaVersion)
    
    'Clients
    mClientCount = mDBFile.ReadLong
    ReDim mClients(mClientCount - 1)
    For a = 0 To mClientCount - 1
        Set mClients(a) = New CClient: mClients(a).Unserialize Me, schemaversion, mDBFile
    Next a
    
    'Appointments
    mAppointmentCount = mDBFile.ReadLong
    ReDim mAppointments(mAppointmentCount - 1)
    For a = 0 To mAppointmentCount - 1
        Set mAppointments(a) = New CAppointment: mAppointments(a).Unserialize Me, schemaversion, mDBFile
    Next a
    
    'ApptCliLinks
    mApptCliLinkCount = mDBFile.ReadLong
    ReDim mApptCliLinks(mApptCliLinkCount - 1)
    For a = 0 To mApptCliLinkCount - 1
        Set mApptCliLinks(a) = New CApptCliLink: mApptCliLinks(a).Unserialize Me, schemaversion, mDBFile
    Next a
    
    'TaxReturns
    mTaxReturnCount = mDBFile.ReadLong
    ReDim mTaxReturns(mTaxReturnCount - 1)
    For a = 0 To mTaxReturnCount - 1
        Set mTaxReturns(a) = New CTaxReturn: mTaxReturns(a).Unserialize Me, schemaversion, mDBFile
    Next a
    
    'Database footer
    If mDBFile.ReadLong <> schemaversion Then Err.Raise 1, , FormatErrorMsg(errDamagedDB, mDBFilename)
    
    '-------------------------

    mIsDirty = False
Else
    If vCreateIfNotExists Then
        'If the database doesn't exist, create it
        CreateFile mDBFilename
        mIsDirty = False
        RaiseEvent OnCreateInitialDB
    Else
        Err.Raise 1, , FormatErrorMsg(errDBFileMissing, mDBFilename)
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

Dim tempfile As String, bakfile As String, a As Long

If Not mLoaded Then Err.Raise 1, , errNotLoaded

If mIsWriteable Then
    If mIsDirty Then
        'Open the file
        Dim mTempFile As CMNMOD_CFileHandler
        tempfile = mDBFilename & ".sav"
        Set mTempFile = OpenFile(tempfile, mBinary_Output)
        
        '-------------------------
        
        'Database header
        mTempFile.WriteLong LatestSchemaVersion
        
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
        mTempFile.WriteLong LatestSchemaVersion
        
        '-------------------------
        
        'Close the files and release any locks
        mTempFile.CloseFile: Set mTempFile = Nothing
        If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing
        
        'Make a backup as of the beginning of the previous day, but don't overwrite it once it has been made
        bakfile = mDBFilename & ".bak" & Format(Date - 1, "yyyymmdd")
        If FileExists(mDBFilename) And Not FileExists(bakfile) Then RenameFile mDBFilename, bakfile, False
        
        'Move the temp file into position as the new database; then reestablish the lock
        RenameFile tempfile, mDBFilename, True
        Set mDBFile = OpenFile(mDBFilename, mBinary_Input)
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

'This function does not raise any errors if the database is not loaded; that way, the calling function need not check first

If mLoaded Then
    'Release the file lock
    If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing

    'Clear out the data in memory
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
'Log of database changes
'#################################################################################

'EHT=None
Public Sub AddLogEntry(logentry As String)

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
'Lookups by primary key
'#################################################################################

'EHT=None
Public Function Client(vClientID As Long) As CClient
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
For a = mClientCount - 1 To 0 Step -1
    If mClients(a).ID = vClientID Then
        Set Client = mClients(a)
        Exit For
    End If
Next a
End Function

'[ApptCliLink omitted]

'EHT=None
Public Function Appointment(vAppointmentID As Long) As CAppointment
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, o As CAppointment
For a = mAppointmentCount - 1 To 0 Step -1
    If mAppointments(a).ID = vAppointmentID Then
        Set Appointment = mAppointments(a)
        Exit For
    End If
Next a
End Function

'EHT=None
Public Function TaxReturn(vClientID As Long, vTaxYear As Long) As CTaxReturn
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, t As CTaxReturn
For a = mTaxReturnCount - 1 To 0 Step -1
    Set t = mTaxReturns(a)
    If t.ClientID = vClientID And t.TaxYear = vTaxYear Then
        Set TaxReturn = t
        Exit For
    End If
Next a
End Function





'#################################################################################
'One-to-many relational lookups
'#################################################################################

'EHT=None
Public Function ApptCliLinksByClient(vClientID As Long) As CApptCliLink()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As CApptCliLink
For a = mApptCliLinkCount - 1 To 0 Step -1
    If mApptCliLinks(a).ClientID = vClientID Then
        ReDim Preserve l(c)
        Set l(c) = mApptCliLinks(a)
        c = c + 1
    End If
Next a
ApptCliLinksByClient = l
End Function

'EHT=None
Public Function ApptCliLinksByAppointment(vAppointmentID As Long) As CApptCliLink()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As CApptCliLink
For a = mApptCliLinkCount - 1 To 0 Step -1
    If mApptCliLinks(a).AppointmentID = vAppointmentID Then
        ReDim Preserve l(c)
        Set l(c) = mApptCliLinks(a)
        c = c + 1
    End If
Next a
ApptCliLinksByAppointment = l
End Function

'EHT=None
Public Function TaxReturnsByClient(vClientID As Long) As CTaxReturn()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As CTaxReturn
For a = mTaxReturnCount - 1 To 0 Step -1
    If mTaxReturns(a).ClientID = vClientID Then
        ReDim Preserve l(c)
        Set l(c) = mTaxReturns(a)
        c = c + 1
    End If
Next a
TaxReturnsByClient = l
End Function





'#################################################################################
'Many-to-many relational lookups
'#################################################################################

'EHT=None
Public Function AppointmentsByClient(vClientID As Long) As CAppointment()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As CAppointment
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .ClientID = vClientID Then
            ReDim Preserve l(c)
            Set l(c) = Appointment(.AppointmentID)
            c = c + 1
        End If
    End With
Next a
AppointmentsByClient = l
End Function

'EHT=None
Public Function ClientsByAppointment(vAppointmentID As Long) As CClient()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As CClient
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .AppointmentID = vAppointmentID Then
            ReDim Preserve l(c)
            Set l(c) = Client(.ClientID)
            c = c + 1
        End If
    End With
Next a
ClientsByAppointment = l
End Function
