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
Private Const errTableMarkerNotFound = "Database file '%1%' has been damaged. Table marker not found at seek position %2%."
Private Const errCountOnDiskMismatch1 = "Database file '%1%' has been damaged. CountOnDisk read from file does not match what was read on initial open of the database."
Private Const errCountOnDiskMismatch2 = "Count mismatch during save. CountInMemory + PendingCountOnDisk <> CountTotal"

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
Private mTaxReturns() As DBModelTaxReturn
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
    CountInMemoryDeleted As Long 'How many of CountInMemory are marked as deleted
    OldestDataLoaded As Long    'This may mean different things for different models (oldest year for Schedules and TaxReturns, lowest ID number for Clients, etc.)
    'For the complete data in the database...
    SeekPosition As Long        'Position in the database file that this table's data begins
    CountOnDisk As Long         'How many model instances exist in the most recent copy of the database file
    CountTotal As Long          'How many model instances really exist, even if not loaded into memory
    NextPK(1 To PKCount) As Long    'Primary key to use for the next model instance added to the database (not used for all models)
    'Temporary variables
    PendingSeekPosition As Long 'Used by SaveChanges and SaveTable
    PendingCountOnDisk As Long  'Used by SaveChanges and SaveTable

    InUse As Boolean            'Temporary, for development purposes only
End Type
Private Const NumberOfTables = 5
Private TableInfo(NumberOfTables - 1) As typeTableInfo
Private Const TableMarker = &HF9F9F900

'Events
Public Event OnFinishedLoading()
Public Event OnDirtyChanged(IsDirty As Boolean)





'#################################################################################
'Load & save functions
'#################################################################################

'EHT=Standard
Public Function ConnectToDatabase(vDBFilename As String, Optional vReadOnly As Boolean, Optional vCreateIfNotExists As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim p As Long, a As Long, tableindex As Long
Dim PK(1 To PKCount) As Long

'Initialize
If mLoaded Then Err.Raise 1, , errAlreadyLoaded
mIsWriteable = Not vReadOnly
mDBFilename = vDBFilename

TableInfo(tClients).InUse = True
TableInfo(tTaxReturns).InUse = True

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

    'Initialize the tables
    For tableindex = 0 To UBound(TableInfo)
        With TableInfo(tableindex)
            .ArrayUbound = -1
            .CountInMemory = 0
            .OldestDataLoaded = MaxLong     'None have been loaded yet
        End With
    Next tableindex

    '-------------------------

    'Database header
    mSchemaVersionOnDisk = mDBFile.ReadLong
    If (mSchemaVersionOnDisk < LowestSupportedSchemaVersion) Or (mSchemaVersionOnDisk > LatestSchemaVersion) Then Err.Raise 1, , FormatErrorMsg(errSchemaNotSupported, mSchemaVersionOnDisk, LowestSupportedSchemaVersion, LatestSchemaVersion)

    'Clients
    LoadTable tClients, MaxLong             'Load nothing for now

    'TaxReturns
    LoadTable tTaxReturns, FileToOpen_Year  'Only load current year

    'Appointments
    mAppointmentCount = mDBFile.ReadLong
    If mAppointmentCount = 0 Then
        Erase mAppointments
    Else
        ReDim mAppointments(mAppointmentCount - 1)
        For a = 0 To mAppointmentCount - 1
            For p = 1 To PKCount
                PK(p) = mDBFile.ReadLong
            Next p
            Set mAppointments(a) = New DBModelAppointment: mAppointments(a).InitFromFile mDBFile, mSchemaVersionOnDisk, PK
        Next a
    End If

    'ApptCliLinks
    mApptCliLinkCount = mDBFile.ReadLong
    If mApptCliLinkCount = 0 Then
        Erase mApptCliLinks
    Else
        ReDim mApptCliLinks(mApptCliLinkCount - 1)
        For a = 0 To mApptCliLinkCount - 1
            For p = 1 To PKCount
                PK(p) = mDBFile.ReadLong
            Next p
            Set mApptCliLinks(a) = New DBModelApptCliLink: mApptCliLinks(a).InitFromFile mDBFile, mSchemaVersionOnDisk, PK
        Next a
    End If

    'Schedules
    mScheduleCount = mDBFile.ReadLong
    If mScheduleCount = 0 Then
        Erase mSchedules
    Else
        ReDim mSchedules(mScheduleCount - 1)
        For a = 0 To mScheduleCount - 1
            For p = 1 To PKCount
                PK(p) = mDBFile.ReadLong
            Next p
            Set mSchedules(a) = New DBModelSchedule: mSchedules(a).InitFromFile mDBFile, mSchemaVersionOnDisk, PK
        Next a
    End If

    'Database footer
    If mDBFile.ReadLong <> mSchemaVersionOnDisk Then Err.Raise 1, , FormatErrorMsg(errDamagedDB, mDBFilename)

    '-------------------------

    ClearDirty
ElseIf vCreateIfNotExists Then
    '<Temporary development code>
    Dim fh As CMNMOD_CFileHandler, t$, s$()
    Set fh = OpenFile(vDBFilename & ".sample", mLineByLine_Input)
    tableindex = -1
    Do Until fh.EndOfFile
        t$ = fh.ReadLine
        If Len(t$) > 0 Then
            s$ = Split(t$, vbTab)
            If Mid$(t$, 1, 1) = "#" Then
                tableindex = s$(1)
                With TableInfo(tableindex)
                    If .InUse Then
                        .CountOnDisk = 0
                        .CountInMemory = s$(2)
                        .CountTotal = .CountInMemory
                        .ArrayUbound = .CountInMemory + 2
                        For p = 1 To PKCount
                            .NextPK(p) = 1
                        Next p
                        .OldestDataLoaded = MinLong
                    End If
                End With
                Select Case tableindex
                Case tClients:                                  ReDim mClients(TableInfo(tableindex).ArrayUbound)
                Case tTaxReturns:                               ReDim mTaxReturns(TableInfo(tableindex).ArrayUbound)
                Case tAppointments: mAppointmentCount = s$(2):  ReDim mAppointments(mAppointmentCount - 1)
                Case tApptCliLinks: mApptCliLinkCount = s$(2):  ReDim mApptCliLinks(mApptCliLinkCount - 1)
                Case tSchedules:    mScheduleCount = s$(2):     ReDim mSchedules(mScheduleCount - 1)
                Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
                End Select
                a = 0
            Else
                For p = 1 To PKCount
                    PK(p) = CLng(s$(p - 1))
                Next p
                With TableInfo(tableindex)
                    If .InUse Then
                        For p = 1 To PKCount
                            If PK(p) >= .NextPK(p) Then .NextPK(p) = PK(p) + 1
                        Next p
                        If a >= .CountInMemory Then Err.Raise 1, , "Too many items in table #" & tableindex
                    End If
                End With
                Select Case tableindex
                Case tClients:      Set mClients(a) = New DBModelClient:            mClients(a).TEMP_InitWithData Me, PK, s$(PKCount), s$(PKCount + 1), s$(PKCount + 2), s$(PKCount + 3)
                Case tTaxReturns:   Set mTaxReturns(a) = New DBModelTaxReturn:      mTaxReturns(a).TEMP_InitWithData Me, PK, CLng(s$(PKCount)), CLng(s$(PKCount + 1))
                Case tAppointments: Set mAppointments(a) = New DBModelAppointment:  mAppointments(a).TEMP_InitWithData Me, PK, CDate(s$(PKCount)), CLng(s$(PKCount + 1)), CLng(s$(PKCount + 2))
                Case tApptCliLinks: Set mApptCliLinks(a) = New DBModelApptCliLink:  mApptCliLinks(a).TEMP_InitWithData Me, PK, CBool(s$(PKCount))
                Case tSchedules:    Set mSchedules(a) = New DBModelSchedule:        mSchedules(a).TEMP_InitWithData Me, PK
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

'EHT=None
Private Sub LoadTable(tableindex As Long, initialextent As Long)
With TableInfo(tableindex)
    .SeekPosition = mDBFile.SeekPosition
    If mDBFile.ReadLong <> (TableMarker + tableindex) Then Err.Raise 1, , FormatErrorMsg(errTableMarkerNotFound, mDBFilename, .SeekPosition)
    .CountOnDisk = mDBFile.ReadLong
    .CountTotal = .CountOnDisk
End With
LoadMoreRecords tableindex, initialextent
End Sub

'EHT=Cleanup2
Private Sub LoadMoreRecords(tableindex As Long, newoldest As Long)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

Dim PK(1 To PKCount) As Long, a&, p&, countofnewloaded As Long
Dim tempobject As Object

With TableInfo(tableindex)
    If Not .InUse Then Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)

    If mLoaded Then
        'Since the initial load process (ConnectToDatabase) has already been completed, we have no requirement
        '  to leave the file pointer at the end of this table. So if there's nothing to load, just exit.
        If newoldest >= .OldestDataLoaded Then Exit Sub
    End If

    'Create a temporary object
    Select Case tableindex
    Case tClients:      Set tempobject = New DBModelClient
    Case tTaxReturns:   Set tempobject = New DBModelTaxReturn
    Case tAppointments: Set tempobject = New DBModelAppointment
    Case tApptCliLinks: Set tempobject = New DBModelApptCliLink
    Case tSchedules:    Set tempobject = New DBModelSchedule
    Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
    End Select

    If mLoaded Then
        'Since the initial load process (ConnectToDatabase) has already been completed, we must do our own
        '  seeking to the table and verifying the marker.
        mDBFile.SeekPosition = .SeekPosition
        If mDBFile.ReadLong <> (TableMarker + tableindex) Then Err.Raise 1, , FormatErrorMsg(errTableMarkerNotFound, mDBFilename, .SeekPosition)
        If mDBFile.ReadLong <> .CountOnDisk Then Err.Raise 1, , FormatErrorMsg(errCountOnDiskMismatch1, mDBFilename)
    End If
    For a = 0 To .CountOnDisk - 1
        For p = 1 To PKCount
            PK(p) = mDBFile.ReadLong
        Next p
        If Not mLoaded Then
            'Only during the initial load process (ConnectToDatabase), we calculate the next primary key available for use.
            For p = 1 To PKCount
                If PK(p) >= .NextPK(p) Then .NextPK(p) = PK(p) + 1
            Next p
        End If
        If (PK(1) >= newoldest) And (PK(1) < .OldestDataLoaded) Then
            'Load this record
            tempobject.InitFromFile mDBFile, mSchemaVersionOnDisk, PK
            Set tempobject.AttachedDatabase = Me
            p = CalcNewIndexAndExpandIfNeeded(tableindex)

            'Store the object, then recreate it so the next InitFromFile doesn't affect the object we just saved
            Select Case tableindex
            Case tClients:      Set mClients(p) = tempobject:       Set tempobject = New DBModelClient
            Case tTaxReturns:   Set mTaxReturns(p) = tempobject:    Set tempobject = New DBModelTaxReturn
            Case tAppointments: Set mAppointments(p) = tempobject:  Set tempobject = New DBModelAppointment
            Case tApptCliLinks: Set mApptCliLinks(p) = tempobject:  Set tempobject = New DBModelApptCliLink
            Case tSchedules:    Set mSchedules(p) = tempobject:     Set tempobject = New DBModelSchedule
            Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
            End Select

            'If InitFromFile was successful, we can increment the count. But keep track of how many we added to the array. If this function
            '  does not completely succeed, we'll have to subtract that amount from the count to rollback the transaction.
            .CountInMemory = .CountInMemory + 1
            countofnewloaded = countofnewloaded + 1
        Else
            'Skip this record, but still InitFromFile, so the file pointer moves properly
            tempobject.InitFromFile mDBFile, mSchemaVersionOnDisk, PK
        End If
    Next a

    'Only once this entire function has succeeded, we may update OldestDataLoaded. Otherwise, an error colud result in data loss.
    .OldestDataLoaded = newoldest
End With

CLEANUP: INCLEANUP = True
    If HASERROR Then
        'Subtract from the count the number of items we added, so we don't end up with duplicate items in the database on next save.
        'We don't need to remove the items themselves, though, since they'll be overwritten eventually.
        If countofnewloaded > 0 Then TableInfo(tableindex).CountInMemory = TableInfo(tableindex).CountInMemory - countofnewloaded
    End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "LoadMoreRecords", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Sub

'EHT=None
Private Function CalcNewIndexAndExpandIfNeeded(tableindex As Long) As Long
With TableInfo(tableindex)
    If .CountInMemory > .ArrayUbound Then
        .ArrayUbound = .ArrayUbound + 3
        Select Case tableindex
        Case tClients:      ReDim Preserve mClients(.ArrayUbound)
        Case tTaxReturns:   ReDim Preserve mTaxReturns(.ArrayUbound)
        Case tAppointments: ReDim Preserve mAppointments(.ArrayUbound)
        Case tApptCliLinks: ReDim Preserve mApptCliLinks(.ArrayUbound)
        Case tSchedules:    ReDim Preserve mSchedules(.ArrayUbound)
        Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
        End Select
    End If
    CalcNewIndexAndExpandIfNeeded = .CountInMemory
End With
End Function

'EHT=Cleanup1
Public Function SaveChanges() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

'Instead of raising an error if the database is not loaded, we return False; that way, the calling function need not check first
'Also note that a return Value of True doesn't mean we saved anything; it just means that everything is saved (even if it was already saved)

Dim tempfile As CMNMOD_CFileHandler, tempfilename As String, bakfilename As String, a As Long, tableindex As Long

If mLoaded And mIsWriteable Then
    If mIsDirty Then
        'Open the file
        tempfilename = mDBFilename & ".sav"
        Set tempfile = OpenFile(tempfilename, mBinary_Output)

        '-------------------------

        'Database header
        tempfile.WriteLong LatestSchemaVersion

        'Clients
        WriteTable tClients, tempfile

        'TaxReturns
        WriteTable tTaxReturns, tempfile

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

        'Now that everything has succeeded, we can update the counts
        For a = 0 To UBound(TableInfo)
            TableInfo(a).CountOnDisk = TableInfo(a).PendingCountOnDisk
            TableInfo(a).SeekPosition = TableInfo(a).PendingSeekPosition
        Next a

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

'EHT=None
Private Sub WriteTable(tableindex As Long, tempfile As CMNMOD_CFileHandler)
Dim a As Long, b As Long, PK(1 To PKCount) As Long, tempobject As Object

'Create a temporary object
Select Case tableindex
Case tClients:      Set tempobject = New DBModelClient
Case tTaxReturns:   Set tempobject = New DBModelTaxReturn
Case tAppointments: Set tempobject = New DBModelAppointment
Case tApptCliLinks: Set tempobject = New DBModelApptCliLink
Case tSchedules:    Set tempobject = New DBModelSchedule
Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
End Select

With TableInfo(tableindex)
    'Before we do any writing on the new database, store the current seek position for future saves
    .PendingCountOnDisk = 0
    .PendingSeekPosition = tempfile.SeekPosition

    'Write the total number of objects that will be written, regardless if they are loaded into memory or not
    tempfile.WriteLong (TableMarker + tableindex)
    tempfile.WriteLong .CountTotal - .CountInMemoryDeleted

    If Not mDBFile Is Nothing Then
        'If we have a lock open to an existing database, seek the old file to where this table began originally
        mDBFile.SeekPosition = .SeekPosition
        If mDBFile.ReadLong <> (TableMarker + tableindex) Then Err.Raise 1, , FormatErrorMsg(errTableMarkerNotFound, mDBFilename, .SeekPosition)
        If mDBFile.ReadLong <> .CountOnDisk Then Err.Raise 1, , FormatErrorMsg(errCountOnDiskMismatch1, mDBFilename)

        'Now scan through every record on-disk
        For a = 0 To .CountOnDisk - 1
            For b = 1 To PKCount
                PK(b) = mDBFile.ReadLong
            Next b
            If PK(1) < .OldestDataLoaded Then
                'If it is older than OldestDataLoaded, then we must copy the data from the old database
                tempobject.InitFromFile mDBFile, mSchemaVersionOnDisk, PK
                tempobject.Serialize tempfile
                .PendingCountOnDisk = .PendingCountOnDisk + 1
            Else
                'Else, we already have it in memory, so skip the on-disk version for now
                tempobject.InitFromFile mDBFile, mSchemaVersionOnDisk, PK
            End If
        Next a
    End If

    'Whether we had an old database or not, the count of objects we copied over plus the number we have left in memory to write must
    'represent the entire dataset without overlap
    If (.CountInMemory + .PendingCountOnDisk) <> .CountTotal Then Err.Raise 1, , errCountOnDiskMismatch2

    'Now store the ones we have in memory, that have not been marked for deletion (those were already excluded from the total count earlier)
    For a = 0 To .CountInMemory - 1
        Select Case tableindex
        Case tClients:      If Not mClients(a).Deleted Then mClients(a).Serialize tempfile
        Case tTaxReturns:   If Not mTaxReturns(a).Deleted Then mTaxReturns(a).Serialize tempfile
        Case tAppointments: If Not mAppointments(a).Deleted Then mAppointments(a).Serialize tempfile
        Case tApptCliLinks: If Not mApptCliLinks(a).Deleted Then mApptCliLinks(a).Serialize tempfile
        Case tSchedules:    If Not mSchedules(a).Deleted Then mSchedules(a).Serialize tempfile
        Case Else:  Err.Raise 1, , FormatErrorMsg(errUnknownTable, tableindex)
        End Select
        .PendingCountOnDisk = .PendingCountOnDisk + 1
    Next a
End With
End Sub

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
'Lookups by primary key, along with the related Add and Delete methods
'#################################################################################

'EHT=None
Public Function FindClient(vClientID As Long, Optional ByRef oFoundIndex As Long) As DBModelClient
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vClientID < 1 Then Exit Function
LoadMoreRecords tClients, vClientID
For a = TableInfo(tClients).CountInMemory - 1 To 0 Step -1
    If mClients(a).ID = vClientID Then
        If Not mClients(a).Deleted Then
            Set FindClient = mClients(a)
            oFoundIndex = a
            Exit For
        End If
    End If
Next a
End Function

'EHT=None
Public Sub DeleteClient(vClientID As Long)
Dim c As DBModelClient, i As Long, a As Long, b As Long
Set c = FindClient(vClientID, i)        'This will check mLoaded for us
If c Is Nothing Then Exit Sub
'Remove foreign key references to this record first
LoadMoreRecords tTaxReturns, MinLong    'Load all into memory
'LoadMoreRecords tApptCliLinks, MinLong  'Load all into memory
With TableInfo(tTaxReturns)
    For b = 0 To .CountInMemory - 1
        If mTaxReturns(b).ClientID = vClientID Then
            If Not mTaxReturns(b).Deleted Then
                mTaxReturns(b).Deleted = True
                .CountInMemoryDeleted = .CountInMemoryDeleted + 1
            End If
        End If
    Next b
End With
'With TableInfo(tApptCliLinks)
'    For b = 0 To .CountInMemory - 1
'        If mApptCliLinks(b).ClientID = vClientID Then
'            If Not mApptCliLinks(b).Deleted Then
'                mApptCliLinks(b).Deleted = True
'                .CountInMemoryDeleted = .CountInMemoryDeleted + 1
'            End If
'        End If
'    Next b
'End With
'Then delete the client
With TableInfo(tClients)
    mClients(i).Deleted = True
    .CountInMemoryDeleted = .CountInMemoryDeleted + 1
End With
SetDirty
End Sub

'EHT=None
Public Sub AddClient(m As DBModelClient)
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
With TableInfo(tClients)
    'Since the primary key is autoincrement, there is no need to check for duplicates. We can trust NextPK.
    m.ID = .NextPK(1)
    'Add the new item to the database
    a = CalcNewIndexAndExpandIfNeeded(tClients)
    Set mClients(a) = m
    'If we assume NextPK(1) is correct, then there cannot be any item with a PK(1) higher than NextPK(1).
    'Therefore, we can (and must) lower OldestDataLoaded to match without actually loading more data, so
    '  our new PK(1) will be >= OldestDataLoaded, to avoid data corruption and duplication
    If .OldestDataLoaded > .NextPK(1) Then .OldestDataLoaded = .NextPK(1)
    'Update counts last
    .CountInMemory = .CountInMemory + 1
    .CountTotal = .CountTotal + 1
    .NextPK(1) = .NextPK(1) + 1
    SetDirty
End With
End Sub

'EHT=None
Public Function FindTaxReturn(vTaxYear As Long, vClientID As Long) As DBModelTaxReturn
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, t As DBModelTaxReturn
If vTaxYear < 1 Or vClientID < 1 Then Exit Function
LoadMoreRecords tTaxReturns, vTaxYear
For a = TableInfo(tTaxReturns).CountInMemory - 1 To 0 Step -1
    Set t = mTaxReturns(a)
    If t.TaxYear = vTaxYear Then
        If t.ClientID = vClientID Then
            If Not t.Deleted Then
                Set FindTaxReturn = t
                Exit For
            End If
        End If
    End If
Next a
End Function

'EHT=None
Public Sub DeleteTaxReturn(vTaxYear As Long, vClientID As Long)
If Not mLoaded Then Err.Raise 1, , errNotLoaded

End Sub

'EHT=None
Public Sub AddTaxReturn(m As DBModelTaxReturn)
If Not mLoaded Then Err.Raise 1, , errNotLoaded
'Check for duplicates (and load records back that old in the process)

'LoadMoreRecords tTaxReturns, c.ID

End Sub

'EHT=None
Public Function FindAppointment(vAppointmentID As Long) As DBModelAppointment
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vAppointmentID < 1 Then Exit Function
For a = mAppointmentCount - 1 To 0 Step -1
    If mAppointments(a).ID = vAppointmentID Then
        If Not mAppointments(a).Deleted Then
            Set FindAppointment = mAppointments(a)
            Exit For
        End If
    End If
Next a
End Function

'EHT=None
Public Function FindSchedule(vCalendarYear As Long) As DBModelSchedule
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long
If vCalendarYear < 1 Then Exit Function
For a = mScheduleCount - 1 To 0 Step -1
    If mSchedules(a).CalendarYear = vCalendarYear Then
        If Not mSchedules(a).Deleted Then
            Set FindSchedule = mSchedules(a)
            Exit For
        End If
    End If
Next a
End Function





'#################################################################################
'Many-to-many relational lookups
'#################################################################################

'EHT=None
Public Function FindClientsByAppointment(vAppointmentID As Long) As DBModelClient()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelClient
If vAppointmentID < 1 Then Exit Function
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .AppointmentID = vAppointmentID Then
            If Not .Deleted Then
                ReDim Preserve l(c)
                Set l(c) = FindClient(.ClientID)
                c = c + 1
            End If
        End If
    End With
Next a
FindClientsByAppointment = l
End Function

'EHT=None
Public Function FindAppointmentsByClient(vClientID As Long) As DBModelAppointment()
If Not mLoaded Then Err.Raise 1, , errNotLoaded
Dim a As Long, c As Long, l() As DBModelAppointment
If vClientID < 1 Then Exit Function
For a = mApptCliLinkCount - 1 To 0 Step -1
    With mApptCliLinks(a)
        If .ClientID = vClientID Then
            If Not .Deleted Then
                ReDim Preserve l(c)
                Set l(c) = FindAppointment(.AppointmentID)
                c = c + 1
            End If
        End If
    End With
Next a
FindAppointmentsByClient = l
End Function
