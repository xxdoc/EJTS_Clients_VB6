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
Private mClients() As DBModelClient
Private mClientCount As Long
Private mAppointments() As DBModelAppointment
Private mAppointmentCount As Long
Private mApptCliLinks() As DBModelApptCliLink
Private mApptCliLinkCount As Long
Private mTaxReturns() As DBModelTaxReturn
Private mTaxReturnCount As Long
Private mSchedules() As DBModelSchedule
Private mScheduleCount As Long

'Events
Public Event OnCreateInitialDB()
Public Event OnFinishedLoading()
Public Event OnDirtyChanged(IsDirty As Boolean)





'#################################################################################
'Load & save functions
'#################################################################################

'EHT=Standard
Public Function ConnectToDatabase(vDBFilename As String, Optional vReadOnly As Boolean, Optional vCreateIfNotExists As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim p As Long, a As Long, loadable As Boolean, schemaversion As Long

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

loadable = FileExists(mDBFilename)
If loadable Then loadable = loadable And (FileLen(mDBFilename) > 0)
If loadable Then
    'File exists, and is not empty, so open it
    Set mDBFile = OpenFile(mDBFilename, mBinary_Input)

    '-------------------------

    'Database header
    schemaversion = mDBFile.ReadLong
    If (schemaversion < LowestSupportedSchemaVersion) Or (schemaversion > LatestSchemaVersion) Then Err.Raise 1, , FormatErrorMsg(errSchemaNotSupported, schemaversion, LowestSupportedSchemaVersion, LatestSchemaVersion)

    'Clients
    mClientCount = mDBFile.ReadLong
    ReDim mClients(mClientCount - 1)
    For a = 0 To mClientCount - 1
        Set mClients(a) = New DBModelClient: mClients(a).Unserialize Me, schemaversion, mDBFile
    Next a

    'Appointments
    mAppointmentCount = mDBFile.ReadLong
    ReDim mAppointments(mAppointmentCount - 1)
    For a = 0 To mAppointmentCount - 1
        Set mAppointments(a) = New DBModelAppointment: mAppointments(a).Unserialize Me, schemaversion, mDBFile
    Next a

    'ApptCliLinks
    mApptCliLinkCount = mDBFile.ReadLong
    ReDim mApptCliLinks(mApptCliLinkCount - 1)
    For a = 0 To mApptCliLinkCount - 1
        Set mApptCliLinks(a) = New DBModelApptCliLink: mApptCliLinks(a).Unserialize Me, schemaversion, mDBFile
    Next a

    'TaxReturns
    mTaxReturnCount = mDBFile.ReadLong
    ReDim mTaxReturns(mTaxReturnCount - 1)
    For a = 0 To mTaxReturnCount - 1
        Set mTaxReturns(a) = New DBModelTaxReturn: mTaxReturns(a).Unserialize Me, schemaversion, mDBFile
    Next a

    'Schedules
    mScheduleCount = mDBFile.ReadLong
    ReDim mSchedules(mScheduleCount - 1)
    For a = 0 To mScheduleCount - 1
        Set mSchedules(a) = New DBModelSchedule: mSchedules(a).Unserialize Me, schemaversion, mDBFile
    Next a

    'Database footer
    If mDBFile.ReadLong <> schemaversion Then Err.Raise 1, , FormatErrorMsg(errDamagedDB, mDBFilename)

    '-------------------------

    ClearDirty
ElseIf vCreateIfNotExists Then
    'If the database doesn't exist, create it
    CreateFile mDBFilename, True
    ClearDirty
    RaiseEvent OnCreateInitialDB
Else
    'If we're not allowed to create it, error
    Err.Raise 1, , FormatErrorMsg(errDBFileMissing, mDBFilename)
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

'This function does not raise any errors if the database is not loaded; that way, the calling function need not check first

Dim tempfile As CMNMOD_CFileHandler, tempfilename As String, bakfilename As String, a As Long

If mLoaded And mIsWriteable Then
    If mIsDirty Then
        'Open the file
        tempfilename = mDBFilename & ".sav"
        Set tempfile = OpenFile(tempfilename, mBinary_Output)

        '-------------------------

        'Database header
        tempfile.WriteLong LatestSchemaVersion

        'Clients
        tempfile.WriteLong mClientCount
        For a = 0 To mClientCount - 1
            mClients(a).Serialize tempfile
        Next a

        'Appointments
        tempfile.WriteLong mAppointmentCount
        For a = 0 To mAppointmentCount - 1
            mAppointments(a).Serialize tempfile
        Next a

        'ApptCliLinks
        tempfile.WriteLong mApptCliLinkCount
        For a = 0 To mApptCliLinkCount - 1
            mApptCliLinks(a).Serialize tempfile
        Next a

        'TaxReturns
        tempfile.WriteLong mTaxReturnCount
        For a = 0 To mTaxReturnCount - 1
            mTaxReturns(a).Serialize tempfile
        Next a

        'Schedules
        tempfile.WriteLong mScheduleCount
        For a = 0 To mScheduleCount - 1
            mSchedules(a).Serialize tempfile
        Next a

        'Database footer
        tempfile.WriteLong LatestSchemaVersion

        '-------------------------

        'Close the files and release any locks
        tempfile.CloseFile: Set tempfile = Nothing
        If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing

        'Make a backup, but don't overwrite it once it has been made. Each backup will contain all changes up to and including the day
        'indicated (.bak20160111 would include all changes from the beginning of time through 1/11 but none of the changes of 1/12).
        'Unfortunately, to make this work, the first backup created will be a 0-byte file (it means that prior to that first save of
        'the day, there was no data).
        If FileExists(mDBFilename) Then
            bakfilename = mDBFilename & ".bak" & Format(Date - 1, "yyyymmdd")
            If Not FileExists(bakfilename) Then RenameFile mDBFilename, bakfilename, False
        End If

        'Move the temp file into position as the new database; then reestablish the lock
        RenameFile tempfilename, mDBFilename, True
        Set mDBFile = OpenFile(mDBFilename, mBinary_Input)
        ClearDirty
    End If

    'Return value of True means that the database file now reflects the data (regardless if we had to write any data to make that happen)
    SaveChanges = True
End If

CLEANUP: INCLEANUP = True
    If Not tempfile Is Nothing Then tempfile.CloseFile: Set tempfile = Nothing

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
    Erase mSchedules
    mScheduleCount = 0
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
RaiseEvent OnDirtyChanged(mIsDirty)
End Sub

'EHT=None
Public Sub ClearDirty()
mIsDirty = False
RaiseEvent OnDirtyChanged(mIsDirty)
End Sub

'EHT=None
Public Property Get IsWriteable() As Boolean
IsWriteable = mIsWriteable
End Property





'#################################################################################
'Lookups by primary key
'#################################################################################

'EHT=None
Public Function Client(vClientID As Long) As DBModelClient
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vClientID < 0 Then Exit Function
For a = mClientCount - 1 To 0 Step -1
    If mClients(a).ID = vClientID Then
        Set Client = mClients(a)
        Exit For
    End If
Next a
End Function

'[ApptCliLink omitted]

'EHT=None
Public Function Appointment(vAppointmentID As Long) As DBModelAppointment
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vAppointmentID < 0 Then Exit Function
For a = mAppointmentCount - 1 To 0 Step -1
    If mAppointments(a).ID = vAppointmentID Then
        Set Appointment = mAppointments(a)
        Exit For
    End If
Next a
End Function

'EHT=None
Public Function TaxReturn(vClientID As Long, vTaxYear As Long) As DBModelTaxReturn
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, t As DBModelTaxReturn
If vClientID < 0 Or vTaxYear < 0 Then Exit Function
For a = mTaxReturnCount - 1 To 0 Step -1
    Set t = mTaxReturns(a)
    If t.ClientID = vClientID And t.TaxYear = vTaxYear Then
        Set TaxReturn = t
        Exit For
    End If
Next a
End Function

'EHT=None
Public Function Schedule(vTaxYear As Long) As DBModelSchedule
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vTaxYear < 0 Then Exit Function
For a = mScheduleCount - 1 To 0 Step -1
    If mSchedules(a).TaxYear = vTaxYear Then
        Set Schedule = mSchedules(a)
        Exit For
    End If
Next a
End Function





'#################################################################################
'One-to-many relational lookups
'#################################################################################

'EHT=None
Public Function ApptCliLinksByClient(vClientID As Long) As DBModelApptCliLink()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelApptCliLink
If vClientID < 0 Then Exit Function
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
Public Function ApptCliLinksByAppointment(vAppointmentID As Long) As DBModelApptCliLink()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelApptCliLink
If vAppointmentID < 0 Then Exit Function
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
Public Function TaxReturnsByClient(vClientID As Long) As DBModelTaxReturn()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelTaxReturn
If vClientID < 0 Then Exit Function
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
Public Function AppointmentsByClient(vClientID As Long) As DBModelAppointment()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelAppointment
If vClientID < 0 Then Exit Function
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
Public Function ClientsByAppointment(vAppointmentID As Long) As DBModelClient()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelClient
If vAppointmentID < 0 Then Exit Function
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
