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
Private Const errDamagedDB = "Database file '%1%' has been damaged. Schema version marker at end of file does not match beginning of file."
Private Const errUnknownTable = "Unknown table index %1%."

'Schema version
Private Const LowestSupportedSchemaVersion As Long = 2
Private Const LatestSchemaVersion As Long = 2

'Database attributes (accessible through class properties)
Private mDBFile As CMNMOD_CFileHandler
Private mDBFilename As String
Private mLogFilename As String
Private mSchemaVersionOnDisk As Long
Private mLoaded As Boolean
Private mIsWriteable As Boolean
Private mIsDirty As Boolean

'Core data
Private mClients() As DBModelClient
Private mClientCount As Long
Private mTaxReturns() As DBModelTaxReturn
Private mTaxReturnCount As Long
Private mAppointments() As DBModelAppointment
Private mAppointmentCount As Long
Private mApptCliLinks() As DBModelApptCliLink
Private mApptCliLinkCount As Long
Private mSchedules() As DBModelSchedule
Private mScheduleCount As Long

'Table data (caching extents, next ID to assign, etc.)
Private Enum enumTable
    tClients
    tTaxReturns
    tAppointments
    tApptCliLinks
    tSchedules
End Enum
Private Type typeTableInfo
    'For the portion of data we've already loaded into memory...
    ArrayUbound As Long         'UBound of the array, even if not every element has been used yet
    CountInMemory As Long       'How many model instances have been cached (exist in the array)
    OldestDataLoaded As Long    'This may mean different things for different models (oldest year for Schedules and TaxReturns, lowest ID number for Clients, etc.)
    'For the complete data in the database...
    SeekPosition As Long        'Position in the database file that this table's data begins
    CountOnDisk As Long         'How many model instances exist in the most recent copy of the database file
    CountTotal As Long          'How many model instances really exist, even if not loaded into memory
    NextPKAvailable As Long     'Primary key to use for the next model instance added to the database (not used for all models)
End Type
Private Const NumberOfTables = 5
Private TableInfo(NumberOfTables - 1) As typeTableInfo

'Events
Public Event OnFinishedLoading()
Public Event OnDirtyChanged(IsDirty As Boolean)





'#################################################################################
'Load & save functions
'#################################################################################

'EHT=Standard
Public Function ConnectToDatabase(vDBFilename As String, Optional vReadOnly As Boolean, Optional vCreateIfNotExists As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim p As Long, a As Long
Dim PK1 As Long, PK2 As Long

'Initialize
If mLoaded Then Err.Raise 1, , errAlreadyLoaded
mIsWriteable = Not vReadOnly
mDBFilename = vDBFilename

'Calculate logfile name
p = InStrRev(vDBFilename, ".")
a = InStrRev(vDBFilename, "\")
If p > a Then   'If there is a path (a > 0), then this also serves the purpose of verifying the . is in the file portion, not the path portion
    'Strip the extension off
    vDBFilename = Left$(vDBFilename, p - 1)
End If
mLogFilename = vDBFilename & ".log"

If FileExists(mDBFilename) Then
    'File exists, and is not empty, so open it
    'We're using a global variable for this, so the file is left OPEN and LOCKED at the end of this function
    Set mDBFile = OpenFile(mDBFilename, mBinary_Input)

    '-------------------------

    'Database header
    mSchemaVersionOnDisk = mDBFile.ReadLong
    If (mSchemaVersionOnDisk < LowestSupportedSchemaVersion) Or (mSchemaVersionOnDisk > LatestSchemaVersion) Then Err.Raise 1, , FormatErrorMsg(errSchemaNotSupported, mSchemaVersionOnDisk, LowestSupportedSchemaVersion, LatestSchemaVersion)

    'Clients
    mClientCount = mDBFile.ReadLong
    If mClientCount = 0 Then
        Erase mClients
    Else
        ReDim mClients(mClientCount - 1)
        For a = 0 To mClientCount - 1
            Set mClients(a) = New DBModelClient: mClients(a).Unserialize Me, mSchemaVersionOnDisk, mDBFile
        Next a
    End If

    'TaxReturns
    mTaxReturnCount = mDBFile.ReadLong
    If mTaxReturnCount = 0 Then
        Erase mTaxReturns
    Else
        ReDim mTaxReturns(mTaxReturnCount - 1)
        For a = 0 To mTaxReturnCount - 1
            Set mTaxReturns(a) = New DBModelTaxReturn: mTaxReturns(a).Unserialize Me, mSchemaVersionOnDisk, mDBFile
        Next a
    End If

    'Appointments
    mAppointmentCount = mDBFile.ReadLong
    If mAppointmentCount = 0 Then
        Erase mAppointments
    Else
        ReDim mAppointments(mAppointmentCount - 1)
        For a = 0 To mAppointmentCount - 1
            Set mAppointments(a) = New DBModelAppointment: mAppointments(a).Unserialize Me, mSchemaVersionOnDisk, mDBFile
        Next a
    End If

    'ApptCliLinks
    mApptCliLinkCount = mDBFile.ReadLong
    If mApptCliLinkCount = 0 Then
        Erase mApptCliLinks
    Else
        ReDim mApptCliLinks(mApptCliLinkCount - 1)
        For a = 0 To mApptCliLinkCount - 1
            Set mApptCliLinks(a) = New DBModelApptCliLink: mApptCliLinks(a).Unserialize Me, mSchemaVersionOnDisk, mDBFile
        Next a
    End If

    'Schedules
    mScheduleCount = mDBFile.ReadLong
    If mScheduleCount = 0 Then
        Erase mSchedules
    Else
        ReDim mSchedules(mScheduleCount - 1)
        For a = 0 To mScheduleCount - 1
            Set mSchedules(a) = New DBModelSchedule: mSchedules(a).Unserialize Me, mSchemaVersionOnDisk, mDBFile
        Next a
    End If

    'Database footer
    If mDBFile.ReadLong <> mSchemaVersionOnDisk Then Err.Raise 1, , FormatErrorMsg(errDamagedDB, mDBFilename)

    '-------------------------

    ClearDirty
ElseIf vCreateIfNotExists Then
    '<Temporary development code>
    Dim fh As CMNMOD_CFileHandler, t$, s$(), tableindex&
    Set fh = OpenFile(vDBFilename & ".sample", mLineByLine_Input)
    tableindex = -1
    Do Until fh.EndOfFile
        t$ = fh.ReadLine
        If Len(t$) > 0 Then
            s$ = Split(t$, vbTab)
            If Mid$(t$, 1, 1) = "#" Then
                tableindex = s$(1)
                Select Case tableindex
                Case tClients:      mClientCount = s$(2):       ReDim mClients(mClientCount - 1)
                Case tTaxReturns:   mTaxReturnCount = s$(2):    ReDim mTaxReturns(mTaxReturnCount - 1)
                Case tAppointments: mAppointmentCount = s$(2):  ReDim mAppointments(mAppointmentCount - 1)
                Case tApptCliLinks: mApptCliLinkCount = s$(2):  ReDim mApptCliLinks(mApptCliLinkCount - 1)
                Case tSchedules:    mScheduleCount = s$(2):     ReDim mSchedules(mScheduleCount - 1)
                Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
                End Select
                a = 0
            Else
                PK1 = CLng(s$(0))
                PK2 = CLng(s$(1))
                Select Case tableindex
                Case tClients:      Set mClients(a) = New DBModelClient:            mClients(a).Init Me, PK1, PK2:      mClients(a).QuickSetData s$(2), s$(3), s$(4), s$(5)
                Case tTaxReturns:   Set mTaxReturns(a) = New DBModelTaxReturn:      mTaxReturns(a).Init Me, PK1, PK2:   mTaxReturns(a).QuickSetData CLng(s$(2)), CLng(s$(3))
                Case tAppointments: Set mAppointments(a) = New DBModelAppointment:  mAppointments(a).Init Me, PK1, PK2: mAppointments(a).QuickSetData CDate(s$(2)), CLng(s$(3)), CLng(s$(4))
                Case tApptCliLinks: Set mApptCliLinks(a) = New DBModelApptCliLink:  mApptCliLinks(a).Init Me, PK1, PK2: mApptCliLinks(a).PrimaryClient = s$(2)
                Case tSchedules:    Set mSchedules(a) = New DBModelSchedule:        mSchedules(a).Init Me, PK1, PK2
                Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
                End Select
                a = a + 1
            End If
        End If
    Loop
    fh.CloseFile
    '</Temporary development code>
    
    SetDirty
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

'Instead of raising an error if the database is not loaded, we return False; that way, the calling function need not check first
'Also note that a return Value of True doesn't mean we saved anything; it just means that everything is saved (even if it was already saved)

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

        'TaxReturns
        tempfile.WriteLong mTaxReturnCount
        For a = 0 To mTaxReturnCount - 1
            mTaxReturns(a).Serialize tempfile
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

        'Schedules
        tempfile.WriteLong mScheduleCount
        For a = 0 To mScheduleCount - 1
            mSchedules(a).Serialize tempfile
        Next a

        'Database footer
        tempfile.WriteLong LatestSchemaVersion

        '-------------------------

        'Close the new file
        tempfile.CloseFile: Set tempfile = Nothing

        'Close the old file to release any locks
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
'A return value of True just means that the database is successfully disconnected (even if it wasn't even connected in the first place)

If mLoaded Then
    'Release the file lock
    If Not mDBFile Is Nothing Then mDBFile.CloseFile: Set mDBFile = Nothing

    'Mark the database unloaded and clear the reference to the file, so if we encounter any error
    'we won't be leaving the database in a state that looks usable
    mLoaded = False
    mDBFilename = ""
    mLogFilename = ""

    'Release the majority of memory being used (no need to erase everything, though, since we use mLoaded to determine if the data is usable)
    Erase mClients
    Erase mTaxReturns
    Erase mAppointments
    Erase mApptCliLinks
    Erase mSchedules
End If

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

'EHT=None
Public Function TaxReturn(vTaxYear As Long, vClientID As Long) As DBModelTaxReturn
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, t As DBModelTaxReturn
If vTaxYear < 0 Or vClientID < 0 Then Exit Function
For a = mTaxReturnCount - 1 To 0 Step -1
    Set t = mTaxReturns(a)
    If t.TaxYear = vTaxYear And t.ClientID = vClientID Then
        Set TaxReturn = t
        Exit For
    End If
Next a
End Function

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

'[ApptCliLink omitted]

'EHT=None
Public Function Schedule(vCalendarYear As Long) As DBModelSchedule
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vCalendarYear < 0 Then Exit Function
For a = mScheduleCount - 1 To 0 Step -1
    If mSchedules(a).CalendarYear = vCalendarYear Then
        Set Schedule = mSchedules(a)
        Exit For
    End If
Next a
End Function





'#################################################################################
'One-to-many relational lookups
'#################################################################################

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





'#################################################################################
'Many-to-many relational lookups
'#################################################################################

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
