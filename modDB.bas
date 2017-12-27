Attribute VB_Name = "modDB"
Option Explicit
Private Const MOD_NAME = "modDB"

'This module should NEVER reference ActiveDBInstance (use LocalDBInstance instead)
'Also, be careful when calling frmmain.SetChangedFlagAndIndication, since that is only relevant to ActiveDBInstance

Public Const CurDBFileSpecVersion = "EJTS-v07"

Public Enum MailingListStatus
    Auto
    EmailOrganizer
    HardCopyOrganizer
    NoOrganizer
End Enum

Public Enum ClientFlags
    'Only 1 of these 2
    PartiallyComplete = 1
    CompletedReturn = 2

    'Only 1 of these 4
    HadAppointment = 4      'Mark this ONLY during posting
    DroppedOff = 8          'Mark this when client drops off (confirmed during posting)
    MailedIn = 16           'Mark this when package received (confirmed during posting)
    NoNeedToFile = 32       'Mark this ONLY during posting

    'Individual flags
    Extension = 64
    IncPtnrTrustEstate = 128    'Cannot by IPTE and NNTF at the same time
    NewClient = 256
    EFiled = 512
    'NoStateReturn = 1024
    ReleasedBeforePayment = 1024    '2048
End Enum
Public Const ClientFlags_DATAITEMUBOUND = 11 - 1

Public Type PersonStruct_v06
    First As String
    Nickname As String
    Initial As String
    Last As String
    DOD As Long
    Phone As String
    Email As String
End Type
Public Type PersonStruct
    First As String
    Nickname As String
    Initial As String
    Last As String
    DOB As Long
    DOD As Long
    Phone As String
    Email As String
End Type
Public Type Client_DBPortion
    ID As Long

    'Person-specific information (see above)
    Person1 As PersonStruct
    Person2 As PersonStruct

    'TaxReturn-specific information
    PhoneHome As String            'Phone numbers must be in 0000000000 or 0000000000x* format
    AddressStreet As String
    AddressCity As String
    AddressState As String
    AddressZipCode As String
    Notes As String
    NumApptSlotsToUse As Long
    Flags As Long
    MailingListStatus As Long       '0-Auto, 1-Email(MLE), 2-HardCopy(MLH), 3-None(NML)

    'Last year's data
    LastYear_MinutesToComplete As Long
    LastYear_PrepFee As Long
    LastYear_Flags As Long
    OldestYearFiled As Long
    NewestYearFiled As Long

    'Posting data
    CompletionDate As Long
    MinutesToComplete As Long
    StateList As String
    PrepFee As Long
    MoneyOwed As Long
    ResultAGI As Long
    ResultFederal As Long
    ResultState As Long

    'Operation notes
    OpNotes As String
End Type
Public Type Client
    c As Client_DBPortion
    'Temp data (do not read or write from database file)
    Temp_RegenerateTempData As Boolean
    Temp_ParsedName As String
    Temp_ApptDate As String
    Temp_ApptPast As Boolean
    Temp_DidntHappen As Boolean
End Type
Public Const Client_DATAITEMUBOUND = 27 - 1     'Excludes temp data

Public Enum AppointmentFlags
    ReminderCall = 1
    Called = 2
    DidntHappen = 4    'True if No Show, same-day cancel, or same-day reschedule
End Enum
Public Const AppointmentFlags_DATAITEMUBOUND = 3 - 1

Public Type Appointment
    ID As Long
    ApptDate As Long
    ApptTimeSlot As Long
    ApptActualTime As Date
    NumSlots As Long
    Flags As Long
    ClientIDs() As Long
    ClientID_Count As Long
    Notes As String
End Type
Public Const Appointment_DATAITEMUBOUND = 8 - 1

Public Const Appointment_NumSlots = 14
Public Const Appointment_NumSlotsUB = Appointment_NumSlots - 1
Public Const Slot_DefaultAccordingToTemplate = -100
Public Const Slot_Available = -1
Public Const Slot_Reserved = -2
Public Const Slot_MealBreak = -3

Public Type ExtraCharge
    ClientName As String
    Description As String
    CompletionDate As Long
    PrepFee As Long
    MoneyOwed As Long
End Type
Public Const ExtraCharge_DATAITEMUBOUND = 5 - 1

Public Type BookkeepingMonth
    CompletionDate As Long
    PrepFee As Long
    MoneyOwed As Long
End Type
Public Const BookkeepingMonth_DATAITEMUBOUND = 3 - 1

Public Type BookkeepingJob
    DisplayName As String
    Months(11) As BookkeepingMonth
End Type
Public Const BookkeepingJob_DATAITEMUBOUND = 13 - 1

Public Type SpecialSearch
    DisplayName As String
    ResultsDisplayMode As String
    SearchString As String
End Type
Public Const SpecialSRCH_DATAITEMUBOUND = 3 - 1

Public Enum enumSettingType
    sStr
    sLng
    sDate
    sBool
    sDbl
End Enum
Public Type Setting
    sName As String
    sType As Long   'enumSettingType
    sValue As Variant
End Type
Public Type Setting_v05
    sName As String
    sValue As Variant
End Type
Public Const SETTING_FIRSTSLOTTIME = "Appointment slots, first time (hours)"
Public Const SETTING_SLOTLENGTH = "Appointment slots, length (minutes)"

Public Type EJTSClientsDB
    Clients() As Client
    Clients_Count As Long

    Appointments() As Appointment
    Appointments_Count As Long

    ApptBitmap() As Long   'CAUTION: Bitmap stores Indexes in program, but IDs in database
    ApptBitmap_StartDate As Long
    ApptBitmap_Count As Long
    Subtitles() As String       'Same length as ApptBitmap

    ExtraCharges() As ExtraCharge
    ExtraCharges_Count As Long

    Bookkeeping() As BookkeepingJob
    Bookkeeping_Count As Long

    SpecialSearches() As SpecialSearch
    SpecialSearches_Count As Long

    Settings() As Setting
    Settings_Count As Long

    Loaded As Boolean
    IsWriteable As Boolean
    Changed As Boolean
    FullPath_DB As String
    FullPath_Log As String
    MakeBakOnNextSave As String
End Type

'EHT=Cleanup1
Function DB_Load(DBFile$, LocalDBInstance As EJTSClientsDB) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If LocalDBInstance.Loaded Then Err.Raise 1, , "A database file is already loaded. Cannot open another until the first one is closed."

Dim filespec$, bakfile$, footer$, a&, b&, c&, d&
Dim TempDBInstance As EJTSClientsDB
Dim fh As CMNMOD_CFileHandler
Dim ttta() As Appointment
Dim ttte() As ExtraCharge
Dim tttb() As BookkeepingJob
Dim ttts() As SpecialSearch
Dim tttse_v05() As Setting_v05
Dim tttse() As Setting

If Not FileExists(DBFile$) Then Err.Raise 1, , "Database file '" & DBFile$ & "' does not exist"

TempDBInstance.FullPath_DB = DBFile$
Set fh = OpenFile(TempDBInstance.FullPath_DB, mBinary_Input)

'FILEHEADER
filespec$ = fh.ReadString(Len(CurDBFileSpecVersion))
Select Case filespec$
Case CurDBFileSpecVersion
    'We're good!!
Case "EJTS-v04", "EJTS-v05", "EJTS-v06"
    'The code below can convert it to the latest format
    ShowInfoMsg "The selected database file was saved in format " & filespec$ & ", but the latest format is " & CurDBFileSpecVersion & "." & vbCrLf & vbCrLf & "The program will attempt to open it anyway. On next save, it will be upgraded to the latest format, and a backup copy will be made with a .bak extension."
    TempDBInstance.MakeBakOnNextSave = Right$(filespec$, 2)
Case Else
    'So old that we are unable to convert it (I'll have to do it manually)
    ShowErrorMsg "The selected database file was saved in format " & filespec$ & ", but the latest format is " & CurDBFileSpecVersion & "." & vbCrLf & vbCrLf & "This file cannot be opened."
    GoTo CLEANUP
End Select

'Clients
TempDBInstance.Clients_Count = fh.ReadLong
If TempDBInstance.Clients_Count = 0 Then
    Erase TempDBInstance.Clients
Else
    ReDim TempDBInstance.Clients(TempDBInstance.Clients_Count - 1)
    For a = 0 To TempDBInstance.Clients_Count - 1
        If filespec$ >= "EJTS-v07" Then
            'Two more Longs were added, so it's now 160 bytes
            Get #fh.FileNum, , TempDBInstance.Clients(a).c
        Else
            'The old Client_DBPortion was 152 bytes, so we must build it manually
            With TempDBInstance.Clients(a).c
                .ID = fh.ReadLong
                .Person1.First = fh.ReadStringPlusLen
                .Person1.Nickname = fh.ReadStringPlusLen
                .Person1.Initial = fh.ReadStringPlusLen
                .Person1.Last = fh.ReadStringPlusLen
                .Person1.DOB = NullLong     'New field
                .Person1.DOD = fh.ReadLong
                .Person1.Phone = fh.ReadStringPlusLen
                .Person1.Email = fh.ReadStringPlusLen
                .Person2.First = fh.ReadStringPlusLen
                .Person2.Nickname = fh.ReadStringPlusLen
                .Person2.Initial = fh.ReadStringPlusLen
                .Person2.Last = fh.ReadStringPlusLen
                .Person2.DOB = NullLong     'New field
                .Person2.DOD = fh.ReadLong
                .Person2.Phone = fh.ReadStringPlusLen
                .Person2.Email = fh.ReadStringPlusLen
                .PhoneHome = fh.ReadStringPlusLen
                .AddressStreet = fh.ReadStringPlusLen
                .AddressCity = fh.ReadStringPlusLen
                .AddressState = fh.ReadStringPlusLen
                .AddressZipCode = fh.ReadStringPlusLen
                .Notes = fh.ReadStringPlusLen
                .NumApptSlotsToUse = fh.ReadLong
                .Flags = fh.ReadLong
                .MailingListStatus = fh.ReadLong
                .LastYear_MinutesToComplete = fh.ReadLong
                .LastYear_PrepFee = fh.ReadLong
                .LastYear_Flags = fh.ReadLong
                .OldestYearFiled = fh.ReadLong
                .NewestYearFiled = fh.ReadLong
                .CompletionDate = fh.ReadLong
                .MinutesToComplete = fh.ReadLong
                .StateList = fh.ReadStringPlusLen
                .PrepFee = fh.ReadLong
                .MoneyOwed = fh.ReadLong
                .ResultAGI = fh.ReadLong
                .ResultFederal = fh.ReadLong
                .ResultState = fh.ReadLong
                .OpNotes = fh.ReadStringPlusLen

                'Also, there is no longer any NoState flag (1024)
                If Flag_IsSet(.Flags, 1024) Then
                    .StateList = ""
                    .Flags = Flag_Remove(.Flags, 1024)
                Else
                    If Flag_IsSet(.Flags, CompletedReturn) And (Not Flag_IsSet(.Flags, IncPtnrTrustEstate)) Then
                        If .ResultState = NullLong Then
                            Err.Raise 1, , "If NoState wasn't set, then there must be some ResultState"
                        Else
                            If Len(.StateList) = 0 Then
                                'If it was blank before, it meant CA only.
                                .StateList = DB_GetSetting(LocalDBInstance, "GLOBAL_DefaultState")
                            'Else
                                'If a state was listed, it meant that state only (no CA) so leave it as-is
                            End If
                        End If
                    End If
                End If
                If Flag_IsSet(.LastYear_Flags, 1024) Then .LastYear_Flags = Flag_Remove(.LastYear_Flags, 1024)

                'The ReleasedBeforePayment flag (2048) shifted down to the NoState position (1024)
                If Flag_IsSet(.Flags, 2048) Then .Flags = Flag_Remove(.Flags, 2048) Or 1024
                If Flag_IsSet(.LastYear_Flags, 2048) Then .LastYear_Flags = Flag_Remove(.LastYear_Flags, 2048) Or 1024
            End With
        End If
    Next a
End If

'Appointments
TempDBInstance.Appointments_Count = fh.ReadLong
If TempDBInstance.Appointments_Count = 0 Then
    Erase TempDBInstance.Appointments
Else
    ReDim ttta(TempDBInstance.Appointments_Count - 1)
    Get #fh.FileNum, , ttta
    TempDBInstance.Appointments = ttta
End If

'Appointment Bitmap
'CAUTION: Indexes are stored in memory, but IDs in database
TempDBInstance.ApptBitmap_StartDate = fh.ReadLong
TempDBInstance.ApptBitmap_Count = fh.ReadLong
If TempDBInstance.ApptBitmap_Count = 0 Then
    Erase TempDBInstance.ApptBitmap
Else
    ReDim TempDBInstance.ApptBitmap(TempDBInstance.ApptBitmap_Count - 1, Appointment_NumSlotsUB)
    For a = 0 To TempDBInstance.ApptBitmap_Count - 1
        For b = 0 To Appointment_NumSlotsUB
            'Convert IDs to Indexes
            c = fh.ReadLong
            If c >= 0 Then
                TempDBInstance.ApptBitmap(a, b) = -1
                For d = 0 To TempDBInstance.Appointments_Count - 1
                    If TempDBInstance.Appointments(d).ID = c Then
                        TempDBInstance.ApptBitmap(a, b) = d
                        Exit For
                    End If
                Next d
            Else
                TempDBInstance.ApptBitmap(a, b) = c
            End If
        Next b
    Next a
End If

'Subtitles
If TempDBInstance.ApptBitmap_Count = 0 Then
    Erase TempDBInstance.Subtitles
Else
    ReDim TempDBInstance.Subtitles(TempDBInstance.ApptBitmap_Count - 1)
    If filespec$ >= "EJTS-v05" Then
        Get #fh.FileNum, , TempDBInstance.Subtitles
    End If
End If

'Extra charges
TempDBInstance.ExtraCharges_Count = fh.ReadLong
If TempDBInstance.ExtraCharges_Count = 0 Then
    Erase TempDBInstance.ExtraCharges
Else
    ReDim ttte(TempDBInstance.ExtraCharges_Count - 1)
    Get #fh.FileNum, , ttte
    TempDBInstance.ExtraCharges = ttte
End If

'Bookkeeping
TempDBInstance.Bookkeeping_Count = fh.ReadLong
If TempDBInstance.Bookkeeping_Count = 0 Then
    Erase TempDBInstance.Bookkeeping
Else
    ReDim tttb(TempDBInstance.Bookkeeping_Count - 1)
    Get #fh.FileNum, , tttb
    TempDBInstance.Bookkeeping = tttb
End If

'Special searches
TempDBInstance.SpecialSearches_Count = fh.ReadLong
If TempDBInstance.SpecialSearches_Count = 0 Then
    Erase TempDBInstance.SpecialSearches
Else
    ReDim ttts(TempDBInstance.SpecialSearches_Count - 1)
    Get #fh.FileNum, , ttts
    TempDBInstance.SpecialSearches = ttts
    If filespec$ < "EJTS-v07" Then
        For a = 0 To TempDBInstance.SpecialSearches_Count - 1
            If TempDBInstance.SpecialSearches(a).SearchString = "OTHERSTATE<>""""" Then
                TempDBInstance.SpecialSearches(a).SearchString = "STATELIST<>"""" STATELIST<>""" & DB_GetSetting(LocalDBInstance, "GLOBAL_DefaultState") & """"
            End If
        Next a
    End If
End If

If filespec$ >= "EJTS-v06" Then
    'Settings
    TempDBInstance.Settings_Count = fh.ReadLong
    If TempDBInstance.Settings_Count = 0 Then
        Erase TempDBInstance.Settings
    Else
        ReDim tttse(TempDBInstance.Settings_Count - 1)
        Get #fh.FileNum, , tttse
        TempDBInstance.Settings = tttse
    End If
Else
    'Settings
    TempDBInstance.Settings_Count = fh.ReadLong
    If TempDBInstance.Settings_Count = 0 Then
        Erase TempDBInstance.Settings
    Else
        ReDim tttse(TempDBInstance.Settings_Count - 1)
        ReDim tttse_v05(TempDBInstance.Settings_Count - 1)
        Get #fh.FileNum, , tttse_v05
        Dim e&
        For e = 0 To TempDBInstance.Settings_Count - 1
            tttse(e).sName = tttse_v05(e).sName
            tttse(e).sValue = tttse_v05(e).sValue

            'Guess at the type
            Select Case VarType(tttse_v05(e).sValue)
            Case 3, 2:    tttse(e).sType = sLng
            Case 5, 4:    tttse(e).sType = sDbl
            Case 8:     tttse(e).sType = sStr
            Case Else
                Err.Raise 1, , "Unknown type"
            End Select

            'Convert the boolean values
            Select Case tttse(e).sName
            Case "SatCheck-LastDayOfTaxSeason", "Statistics-RememberSelection-0", "Statistics-RememberSelection-1"
                tttse(e).sType = sBool
                If tttse(e).sValue = 1 Then
                    tttse(e).sValue = True
                Else
                    tttse(e).sValue = False
                End If
            End Select

            'Rename several settings
            If tttse(e).sName Like "SatCheck-Txt*" Or tttse(e).sName = "MailingList-PaperSize" Or tttse(e).sName Like "Statistics-RememberSelection-*" Or tttse(e).sName Like "Statistics-LastView-*" Or tttse(e).sName = "SatCheck-LastDayOfTaxSeason" Then
                tttse(e).sName = "_" & tttse(e).sName
            ElseIf tttse(e).sName = "NewClientFeeThreshold" Then
                tttse(e).sName = "Prep fee threshold - new client SAF"
            ElseIf tttse(e).sName = "MailingList-PrepFeeThreshold" Then
                tttse(e).sName = "Prep fee threshold - receive organizer"
            ElseIf tttse(e).sName Like "Statistics-Bell*" Then
                Dim f&
                f = Val(Mid$(tttse(e).sName, 29))
                tttse(e).sName = "Bell curve for statistics tab, range " & ((f \ 2) + 1) & Choose((f Mod 2) + 1, " from", " to")
            Else
                Err.Raise 1, , "Unknown setting '" & tttse(e).sName & "'"
            End If
        Next e
        TempDBInstance.Settings = tttse
    End If
End If

'FILEFOOTER
footer$ = fh.ReadString(Len(CurDBFileSpecVersion))
If footer$ <> filespec$ Then Err.Raise 1, , "File footer does not match file header. Unable to read complete file."

'Now that reading is complete, copy temporary lists to master lists
TempDBInstance.Loaded = True
TempDBInstance.IsWriteable = False
TempDBInstance.FullPath_Log = DB_GenerateLogfileName(DBFile$)
TempDBInstance.Changed = False
LocalDBInstance = TempDBInstance    ' This does copy the actual data, rather than copy the pointer to the data as in other languages

DB_Load = True

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_Load", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Cleanup1
Function DB_Save(LocalDBInstance As EJTSClientsDB) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If Not LocalDBInstance.IsWriteable Then
    'This should never actually happen, since there are other protections in the code
    DB_Save = True  'Pretend to succeed so the user isn't stuck with the window on the screen
    Err.Raise 1, , "Database has been opened in read-only mode, yet DB_Save has been called! Your changes will not be saved."
End If

If Not LocalDBInstance.Loaded Then Err.Raise 1, , "Database file not loaded"

Dim a&, b&, c&, destfile$, tempfile$
Dim fh As CMNMOD_CFileHandler

destfile$ = LocalDBInstance.FullPath_DB

If Len(LocalDBInstance.MakeBakOnNextSave) > 0 Then
    FileCopy destfile$, destfile$ & ".bak" & LocalDBInstance.MakeBakOnNextSave
    LocalDBInstance.MakeBakOnNextSave = ""
End If

tempfile$ = destfile$ & ".sav"
Set fh = OpenFile(tempfile$, mBinary_Output)

'FILEHEADER
fh.WriteString CurDBFileSpecVersion

'Clients
fh.WriteLong LocalDBInstance.Clients_Count
For a = 0 To LocalDBInstance.Clients_Count - 1
    Put #fh.FileNum, , LocalDBInstance.Clients(a).c
Next a

'Appointments
fh.WriteLong LocalDBInstance.Appointments_Count
If LocalDBInstance.Appointments_Count > 0 Then
    Put #fh.FileNum, , LocalDBInstance.Appointments
End If

'Appointment Bitmap
'CAUTION: Indexes are stored in memory, but IDs in database
fh.WriteLong LocalDBInstance.ApptBitmap_StartDate
fh.WriteLong LocalDBInstance.ApptBitmap_Count
For a = 0 To LocalDBInstance.ApptBitmap_Count - 1
    For b = 0 To Appointment_NumSlotsUB
        'Convert Indexes to IDs
        c = LocalDBInstance.ApptBitmap(a, b)
        If c >= 0 Then c = LocalDBInstance.Appointments(c).ID
        fh.WriteLong c
    Next b
Next a
'Day subtitles (all-day appointments, basically)
Put #fh.FileNum, , LocalDBInstance.Subtitles

'Extra charges
fh.WriteLong LocalDBInstance.ExtraCharges_Count
If LocalDBInstance.ExtraCharges_Count > 0 Then
    Put #fh.FileNum, , LocalDBInstance.ExtraCharges
End If

'Bookkeeping
fh.WriteLong LocalDBInstance.Bookkeeping_Count
If LocalDBInstance.Bookkeeping_Count > 0 Then
    Put #fh.FileNum, , LocalDBInstance.Bookkeeping
End If

'Special searches
fh.WriteLong LocalDBInstance.SpecialSearches_Count
If LocalDBInstance.SpecialSearches_Count > 0 Then
    Put #fh.FileNum, , LocalDBInstance.SpecialSearches
End If

'Settings
fh.WriteLong LocalDBInstance.Settings_Count
If LocalDBInstance.Settings_Count > 0 Then
    Put #fh.FileNum, , LocalDBInstance.Settings
End If

'FILEFOOTER
fh.WriteString CurDBFileSpecVersion

fh.CloseFile: Set fh = Nothing
RenameFile tempfile$, destfile$, True
LocalDBInstance.Changed = False

DB_Save = True

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_Save", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Standard
Function DB_GetSetting(LocalDBInstance As EJTSClientsDB, ByVal n$, Optional FormatForScreen As Boolean = False, Optional DontCallSetChangedFlag As Boolean) As Variant
On Error GoTo ERR_HANDLER

Dim a&, nl$
nl$ = LCase$(Trim$(n$))
If Left$(nl$, 7) = "global_" Then
    If Not GSLoaded Then Err.Raise 1, , "Global settings not loaded"
    For a = 0 To GlobalSettings_Count - 1
        If LCase$(GlobalSettings(a).sName) = nl$ Then
            'Found
            If FormatForScreen Then
                DB_GetSetting = DB_FormatSettingForScreen(GlobalSettings(a))
            Else
                DB_GetSetting = GlobalSettings(a).sValue
            End If
            Exit Function
        End If
    Next a
Else
    For a = 0 To LocalDBInstance.Settings_Count - 1
        If LCase$(LocalDBInstance.Settings(a).sName) = nl$ Then
            'Found
            If FormatForScreen Then
                DB_GetSetting = DB_FormatSettingForScreen(LocalDBInstance.Settings(a))
            Else
                DB_GetSetting = LocalDBInstance.Settings(a).sValue
            End If
            Exit Function
        End If
    Next a
End If

'If we get to this point, then we didn't find an existing setting match. So create a new name/value pair with default value
If DB_SetDefaultSettingValue(LocalDBInstance, n$, DontCallSetChangedFlag) Then
    DB_GetSetting = DB_GetSetting(LocalDBInstance, n$, , DontCallSetChangedFlag)
Else
    Err.Raise 1, , "No setting found with name '" & n$ & "' and no default value available to create it."
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_GetSetting", Err
End Function

'EHT=Standard
Sub DB_SetSetting(LocalDBInstance As EJTSClientsDB, ByVal n$, v As Variant, Optional CreateAsTypeIfNone As enumSettingType, Optional DontCallSetChangedFlag As Boolean)
On Error GoTo ERR_HANDLER

Dim a&, nl$, g As Boolean, s As Setting, found As Boolean
nl$ = LCase$(Trim$(n$))
g = Left$(nl$, 7) = "global_"
If g Then
    If Not GSLoaded Then Err.Raise 1, , "Global settings not loaded"
    For a = 0 To GlobalSettings_Count - 1
        If LCase$(GlobalSettings(a).sName) = nl$ Then
            s = GlobalSettings(a)
            found = True
            Exit For    'The variable 'a' will be the slot to update
        End If
    Next a
Else
    For a = 0 To LocalDBInstance.Settings_Count - 1
        If LCase$(LocalDBInstance.Settings(a).sName) = nl$ Then
            s = LocalDBInstance.Settings(a)
            found = True
            Exit For    'The variable 'a' will be the slot to update
        End If
    Next a
End If

If Not found Then
    'Set initial values, because this is a new Setting
    s.sName = n$
    If IsMissing(CreateAsTypeIfNone) Then
        s.sType = sStr
    Else
        s.sType = CreateAsTypeIfNone
    End If
End If

Select Case s.sType
Case sStr:
    s.sValue = CStr(v)
Case sLng:
    If VarType(v) = vbString Then
        If Trim$(LCase$(v)) = "null" Then
            s.sValue = NullLong
        Else
            s.sValue = CLng(Val(Trim$(Replace$(Replace$(v, "$", ""), ",", ""))))
        End If
    Else
        s.sValue = CLng(v)
    End If
Case sDbl:
    If VarType(v) = vbString Then
        If Trim$(LCase$(v)) = "null" Then
            s.sValue = NullDouble
        Else
            s.sValue = Val(Trim$(Replace$(Replace$(v, "$", ""), ",", "")))   'Val already returns a Double
        End If
    Else
        s.sValue = CDbl(v)
    End If
Case sDate:
    If VarType(v) = vbString Then
        If IsDate(Trim$(v)) Then
            s.sValue = CDate(Trim$(v))
        Else
            s.sValue = 0
        End If
    Else
        s.sValue = CDate(v)
    End If
Case sBool:
    If VarType(v) = vbString Then
        Select Case LCase$(Trim$(v))
        Case "t", "true", "y", "yes", "1"
            s.sValue = True
        Case Else
            s.sValue = False
        End Select
    Else
        s.sValue = CBool(v)
    End If
End Select

If found Then
    'Update
    If g Then
        If GlobalSettings(a).sName <> s.sName Or GlobalSettings(a).sType <> s.sType Or GlobalSettings(a).sValue <> s.sValue Then
            GSChanged = True
        End If
        GlobalSettings(a) = s
    Else
        'If we're read-only, then the setting will only persist in the current session
        '  This is a better option than preventing settings from even being changed
        If (VarPtr(ActiveDBInstance) = VarPtr(LocalDBInstance)) And (Not DontCallSetChangedFlag) And LocalDBInstance.IsWriteable Then
            If LocalDBInstance.Settings(a).sName <> s.sName Or LocalDBInstance.Settings(a).sType <> s.sType Or LocalDBInstance.Settings(a).sValue <> s.sValue Then
                frmMain.SetChangedFlagAndIndication     ' CAUTION: this interacts with ActiveDBInstance instead of LocalDBInstance
            End If
        End If
        LocalDBInstance.Settings(a) = s
    End If
Else
    'Add new
    If g Then
        ReDim Preserve GlobalSettings(GlobalSettings_Count)
        GlobalSettings(GlobalSettings_Count) = s
        GlobalSettings_Count = GlobalSettings_Count + 1
        GSChanged = True
    Else
        ReDim Preserve LocalDBInstance.Settings(LocalDBInstance.Settings_Count)
        LocalDBInstance.Settings(LocalDBInstance.Settings_Count) = s
        LocalDBInstance.Settings_Count = LocalDBInstance.Settings_Count + 1
        'If we're read-only, then the setting will only persist in the current session
        '  This is a better option than preventing settings from even being changed
        If (VarPtr(ActiveDBInstance) = VarPtr(LocalDBInstance)) And (Not DontCallSetChangedFlag) And LocalDBInstance.IsWriteable Then
            frmMain.SetChangedFlagAndIndication         ' CAUTION: this interacts with ActiveDBInstance instead of LocalDBInstance
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SetSetting", Err
End Sub

'EHT=Standard
Function DB_SetDefaultSettingValue(LocalDBInstance As EJTSClientsDB, n$, Optional DontCallSetChangedFlag As Boolean) As Boolean
On Error GoTo ERR_HANDLER

DB_SetDefaultSettingValue = True
If n$ = "GLOBAL_DataFolder" Then
    DB_SetSetting LocalDBInstance, n$, "", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_TabOrder_ApptEdit" Then
    DB_SetSetting LocalDBInstance, n$, "txtField,0|txtField,1|txtField,2|txtField,3|txtField,4|btnSave|btnCancel|lstClients|btnMoveUp|btnMoveDown|btnAdd|btnRemove", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_TabOrder_BookkeepingEdit" Then
    DB_SetSetting LocalDBInstance, n$, "txtField,0|txtField,1|txtField,2|btnSave|btnCancel", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_TabOrder_ExtraChargeEdit" Then
    DB_SetSetting LocalDBInstance, n$, "txtField,0|txtField,1|txtField,2|txtField,3|txtField,4|btnSave|btnCancel", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_TabOrder_ClientPost" Then
    DB_SetSetting LocalDBInstance, n$, "optInType,0|optInType,1|optInType,2|optInType,3|txtField,124|txtField,103|txtField,100|txtField,102|txtField,107|txtField,109|txtField,104|txtField,119|txtField,111|txtField,105|txtField,112|txtField,123|txtField,110|txtField,114|txtField,115|txtField,116|txtField,117|txtField,126|txtField,127|txtField,101|txtField,108|txtField,128|txtField,129|txtField,134|txtField,131|chkEFile|txtField,130|txtField,106|txtField,113|btnSavePost|btnCancel", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_TabOrder_ClientEdit" Then
    DB_SetSetting LocalDBInstance, n$, "txtField,100|txtField,101|txtField,102|txtField,103|txtField,105|txtField,135|txtField,106|txtField,104|txtField,107|txtField,108|txtField,109|txtField,110|txtField,112|txtField,136|txtField,113|txtField,111|txtField,114|txtField,115|txtField,116|txtField,117|txtField,118|txtField,119|txtField,120|txtField,121|txtField,122|txtField,123|txtField,124|txtField,126|txtField,127|txtField,128|txtField,129|txtField,130|txtField,131|txtField,132|txtField,133|btnSavePost|btnCancel", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_SearchSyntax_Fields" Then
    DB_SetSetting LocalDBInstance, n$, "id|name|lname,lastname,ln|fname,firstname,fn|ph,phone|street|city|state|zip,zipcode|email|notes|slots|flags|lymin|lyfee|lyflags,lyf|compdate,date|min|statelist|fee|chg,charge|agi|fedresult|stateresult|opnotes", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_SearchSyntax_Flags" Then
    DB_SetSetting LocalDBInstance, n$, "i,in,inc|c,co,cmp,comp|a,ap,appt|d,do|m,mi|nf,nntf|e,x,ex,ext|ipte,ipts|n,nn,new|ef|rbp,rel", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_DefaultState" Then
    DB_SetSetting LocalDBInstance, n$, "CA", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_LocalAreaCode" Then
    DB_SetSetting LocalDBInstance, n$, "909", sStr, DontCallSetChangedFlag
ElseIf n$ = "GLOBAL_PullFilesWeekdaysToSkip" Then
    DB_SetSetting LocalDBInstance, n$, "Sun", sStr, DontCallSetChangedFlag
'------------------------------
ElseIf n$ = "Reminder call if appt scheduled more than" Then
    DB_SetSetting LocalDBInstance, n$, 30, sLng, DontCallSetChangedFlag
ElseIf n$ Like "Schedule Template ?? (*)" Then
    DB_SetSetting LocalDBInstance, n$, String$(Appointment_NumSlots, "A"), sStr, DontCallSetChangedFlag
ElseIf n$ = "Schedule Template B starting date" Then
    DB_SetSetting LocalDBInstance, n$, DateSerial(Year(Date), 1, 15), sDate, DontCallSetChangedFlag
ElseIf n$ = "Schedule Template C starting date" Then
    DB_SetSetting LocalDBInstance, n$, DateSerial(Year(Date), 4, 15), sDate, DontCallSetChangedFlag
ElseIf n$ Like "_SatCheck-Txt*" Then
    DB_SetSetting LocalDBInstance, n$, 0, sLng, DontCallSetChangedFlag
ElseIf n$ = "_SatCheck-LastDayOfTaxSeason" Then
    DB_SetSetting LocalDBInstance, n$, False, sBool, DontCallSetChangedFlag
ElseIf n$ = "Prep fee threshold - receive organizer" Then
    DB_SetSetting LocalDBInstance, n$, 0, sLng, DontCallSetChangedFlag
ElseIf n$ = "Prep fee threshold - new client SAF" Then
    DB_SetSetting LocalDBInstance, n$, 0, sLng, DontCallSetChangedFlag
ElseIf n$ = "_MailingList-PaperSize" Then
    DB_SetSetting LocalDBInstance, n$, 0, sLng, DontCallSetChangedFlag
ElseIf n$ Like "_Statistics-RememberSelection-*" Then
    DB_SetSetting LocalDBInstance, n$, False, sBool, DontCallSetChangedFlag
ElseIf n$ Like "_Statistics-LastView-*" Then
    DB_SetSetting LocalDBInstance, n$, "", sStr, DontCallSetChangedFlag
ElseIf n$ Like "Bell curve for statistics tab, range * *" Then
    DB_SetSetting LocalDBInstance, n$, 0, sLng, DontCallSetChangedFlag
ElseIf n$ = SETTING_FIRSTSLOTTIME Then
    DB_SetSetting LocalDBInstance, n$, 9, sDbl, DontCallSetChangedFlag
ElseIf n$ = SETTING_SLOTLENGTH Then
    DB_SetSetting LocalDBInstance, n$, 45, sLng, DontCallSetChangedFlag
Else
    DB_SetDefaultSettingValue = False
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SetDefaultSettingValue", Err
End Function

'EHT=None
Function DB_FormatSettingForScreen(s As Setting) As String
With s
    Select Case .sType
    Case sDate:
        DB_FormatSettingForScreen = Format(.sValue, "m/d/yyyy")
    Case sLng:
        If .sValue = NullLong Then
            DB_FormatSettingForScreen = "null"
        Else
            DB_FormatSettingForScreen = .sValue
        End If
    Case sDbl:
        If .sValue = NullDouble Then
            DB_FormatSettingForScreen = "null"
        Else
            DB_FormatSettingForScreen = .sValue
        End If
    Case sStr, sBool:
        DB_FormatSettingForScreen = .sValue
    End Select
End With
End Function

'EHT=Standard
Function DB_GetNewClientID(LocalDBInstance As EJTSClientsDB) As Long
On Error GoTo ERR_HANDLER

Dim a&, hid&
For a = 0 To LocalDBInstance.Clients_Count - 1
    If LocalDBInstance.Clients(a).c.ID > hid Then hid = LocalDBInstance.Clients(a).c.ID
Next a
DB_GetNewClientID = hid + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_GetNewClientID", Err
End Function

'EHT=Standard
Function DB_AddClient(LocalDBInstance As EJTSClientsDB, c As Client) As Long
On Error GoTo ERR_HANDLER

' Returns     : Index of new item
ReDim Preserve LocalDBInstance.Clients(LocalDBInstance.Clients_Count)
LocalDBInstance.Clients(LocalDBInstance.Clients_Count) = c
DB_AddClient = LocalDBInstance.Clients_Count
LocalDBInstance.Clients_Count = LocalDBInstance.Clients_Count + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_AddClient", Err
End Function

'EHT=Standard
Function DB_FindClientIndex&(LocalDBInstance As EJTSClientsDB, ID&)
On Error GoTo ERR_HANDLER

Dim a&
For a = 0 To LocalDBInstance.Clients_Count - 1
    If LocalDBInstance.Clients(a).c.ID = ID Then
        DB_FindClientIndex = a
        Exit Function
    End If
Next a
DB_FindClientIndex = -1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_FindClientIndex", Err
End Function

'EHT=Standard
Function DB_GetClientAppt(LocalDBInstance As EJTSClientsDB, cID&, founddate As Date) As Long
On Error GoTo ERR_HANDLER

'Returns index of client's appointment
'If any are today or future, returns closest to today
'Otherwise, returns closest past appointment to today

Dim a&, c&, d As Date, cdt As Date
cdt = Now
Dim fi&, fd As Date, f As Boolean
Dim pi&, pd As Date
pi = -1
For a = 0 To LocalDBInstance.Appointments_Count - 1
    For c = 0 To LocalDBInstance.Appointments(a).ClientID_Count - 1
        If LocalDBInstance.Appointments(a).ClientIDs(c) = cID Then
            d = LocalDBInstance.Appointments(a).ApptDate + LocalDBInstance.Appointments(a).ApptActualTime
            If d < cdt Then
                If d > pd Then
                    pd = d
                    pi = a
                End If
            Else
                If (Not f) Or (d < fd) Then
                    fd = d
                    fi = a
                    f = True
                End If
            End If
            Exit For
        End If
    Next c
Next a
If f Then
    DB_GetClientAppt = fi
    founddate = fd
Else
    If pi < 0 Then
        DB_GetClientAppt = -1
    Else
        DB_GetClientAppt = pi
        founddate = pd
    End If
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_GetClientAppt", Err
End Function

'EHT=Standard
Function DB_AddAppointment(LocalDBInstance As EJTSClientsDB, a As Appointment) As Long
On Error GoTo ERR_HANDLER

' Returns     : Index of new item

ReDim Preserve LocalDBInstance.Appointments(LocalDBInstance.Appointments_Count)
LocalDBInstance.Appointments(LocalDBInstance.Appointments_Count) = a
DB_AddAppointment = LocalDBInstance.Appointments_Count
LocalDBInstance.Appointments_Count = LocalDBInstance.Appointments_Count + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_AddAppointment", Err
End Function

'EHT=Standard
Sub DB_RemoveAppointment(LocalDBInstance As EJTSClientsDB, aindex&)
On Error GoTo ERR_HANDLER

' Do NOT call this procedure from within a With block!!!
Dim a&
For a = aindex To LocalDBInstance.Appointments_Count - 2
    LocalDBInstance.Appointments(a) = LocalDBInstance.Appointments(a + 1)
Next a
LocalDBInstance.Appointments_Count = LocalDBInstance.Appointments_Count - 1
If LocalDBInstance.Appointments_Count = 0 Then
    Erase LocalDBInstance.Appointments
Else
    ReDim Preserve LocalDBInstance.Appointments(LocalDBInstance.Appointments_Count - 1)
End If

'Update bitmap
Dim b&
For a = 0 To LocalDBInstance.ApptBitmap_Count - 1
    For b = 0 To Appointment_NumSlotsUB
        Select Case LocalDBInstance.ApptBitmap(a, b)
        Case aindex
            LocalDBInstance.ApptBitmap(a, b) = Slot_Available
        Case Is > aindex
            LocalDBInstance.ApptBitmap(a, b) = LocalDBInstance.ApptBitmap(a, b) - 1
        End Select
    Next b
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_RemoveAppointment", Err
End Sub

'EHT=Standard
Function DB_GetNewAppointmentID(LocalDBInstance As EJTSClientsDB) As Long
On Error GoTo ERR_HANDLER

Dim a&, hid&
For a = 0 To LocalDBInstance.Appointments_Count - 1
    If LocalDBInstance.Appointments(a).ID > hid Then hid = LocalDBInstance.Appointments(a).ID
Next a
DB_GetNewAppointmentID = hid + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_GetNewAppointmentID", Err
End Function

'EHT=Standard
Function DB_FindAppointmentIndex&(LocalDBInstance As EJTSClientsDB, ID&)
On Error GoTo ERR_HANDLER

Dim a&
For a = 0 To LocalDBInstance.Appointments_Count - 1
    If LocalDBInstance.Appointments(a).ID = ID Then
        DB_FindAppointmentIndex& = a
        Exit Function
    End If
Next a
DB_FindAppointmentIndex& = -1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_FindAppointmentIndex", Err
End Function

'EHT=Standard
Function DB_AddBookkeepingJob(LocalDBInstance As EJTSClientsDB, bk As BookkeepingJob, Optional BeforeIndex As Long = -1) As Long
On Error GoTo ERR_HANDLER

' Returns     : Index of new item
Dim a&
ReDim Preserve LocalDBInstance.Bookkeeping(LocalDBInstance.Bookkeeping_Count)
If BeforeIndex = -1 Then
    LocalDBInstance.Bookkeeping(LocalDBInstance.Bookkeeping_Count) = bk
    DB_AddBookkeepingJob = LocalDBInstance.Bookkeeping_Count
Else
    For a = LocalDBInstance.Bookkeeping_Count - 1 To BeforeIndex Step -1
        LocalDBInstance.Bookkeeping(a + 1) = LocalDBInstance.Bookkeeping(a)
    Next a
    LocalDBInstance.Bookkeeping(BeforeIndex) = bk
    DB_AddBookkeepingJob = BeforeIndex
End If
LocalDBInstance.Bookkeeping_Count = LocalDBInstance.Bookkeeping_Count + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_AddBookkeepingJob", Err
End Function

'EHT=Standard
Sub DB_RemoveBookkeepingJob(LocalDBInstance As EJTSClientsDB, bkindex&)
On Error GoTo ERR_HANDLER

' Do NOT call this procedure from within a With block!!!
Dim a&
For a = bkindex To LocalDBInstance.Bookkeeping_Count - 2
    LocalDBInstance.Bookkeeping(a) = LocalDBInstance.Bookkeeping(a + 1)
Next a
LocalDBInstance.Bookkeeping_Count = LocalDBInstance.Bookkeeping_Count - 1
If LocalDBInstance.Bookkeeping_Count = 0 Then
    Erase LocalDBInstance.Bookkeeping
Else
    ReDim Preserve LocalDBInstance.Bookkeeping(LocalDBInstance.Bookkeeping_Count - 1)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_RemoveBookkeepingJob", Err
End Sub

'EHT=Standard
Function DB_AddExtraCharge(LocalDBInstance As EJTSClientsDB, e As ExtraCharge) As Long
On Error GoTo ERR_HANDLER

' Returns     : Index of new item
ReDim Preserve LocalDBInstance.ExtraCharges(LocalDBInstance.ExtraCharges_Count)
LocalDBInstance.ExtraCharges(LocalDBInstance.ExtraCharges_Count) = e
DB_AddExtraCharge = LocalDBInstance.ExtraCharges_Count
LocalDBInstance.ExtraCharges_Count = LocalDBInstance.ExtraCharges_Count + 1

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_AddExtraCharge", Err
End Function

'EHT=Standard
Sub DB_RemoveExtraCharge(LocalDBInstance As EJTSClientsDB, eindex&)
On Error GoTo ERR_HANDLER

' Do NOT call this procedure from within a With block!!!
Dim a&
For a = eindex To LocalDBInstance.ExtraCharges_Count - 2
    LocalDBInstance.ExtraCharges(a) = LocalDBInstance.ExtraCharges(a + 1)
Next a
LocalDBInstance.ExtraCharges_Count = LocalDBInstance.ExtraCharges_Count - 1
If LocalDBInstance.ExtraCharges_Count = 0 Then
    Erase LocalDBInstance.ExtraCharges
Else
    ReDim Preserve LocalDBInstance.ExtraCharges(LocalDBInstance.ExtraCharges_Count - 1)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_RemoveExtraCharge", Err
End Sub

'EHT=Standard
Sub DB_ClearAndRebuildApptBitmap(LocalDBInstance As EJTSClientsDB)
On Error GoTo ERR_HANDLER

Dim a&, b&, i&, ub&
'Create the default availability
For a = 0 To LocalDBInstance.ApptBitmap_Count - 1
    For b = 0 To Appointment_NumSlotsUB
        LocalDBInstance.ApptBitmap(a, b) = Slot_DefaultAccordingToTemplate
    Next b
Next a
'Fill in where appointments exist
For a = 0 To LocalDBInstance.Appointments_Count - 1
    With LocalDBInstance.Appointments(a)
        i = .ApptDate - LocalDBInstance.ApptBitmap_StartDate
        ub = .ApptTimeSlot + .NumSlots - 1
        If ub > Appointment_NumSlotsUB Then ub = Appointment_NumSlotsUB
        For b = .ApptTimeSlot To ub
            LocalDBInstance.ApptBitmap(i, b) = a
        Next b
    End With
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_ClearAndRebuildApptBitmap", Err
End Sub

'EHT=Standard
Function DB_DayWithinBitmapRange(LocalDBInstance As EJTSClientsDB, Day&) As Boolean
On Error GoTo ERR_HANDLER

Dim a&
a = Day - LocalDBInstance.ApptBitmap_StartDate
If (a >= 0) And (a < LocalDBInstance.ApptBitmap_Count) Then
    DB_DayWithinBitmapRange = True
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_DayWithinBitmapRange", Err
End Function

'EHT=Standard
Sub DB_SlotsClear(LocalDBInstance As EJTSClientsDB, Day&, TimeSlot&, NumSlots&)
On Error GoTo ERR_HANDLER

Dim a&, ub&
If Not DB_DayWithinBitmapRange(LocalDBInstance, Day) Then Exit Sub
ub = TimeSlot + NumSlots - 1
If ub > Appointment_NumSlotsUB Then ub = Appointment_NumSlotsUB
For a = TimeSlot To ub
    LocalDBInstance.ApptBitmap(Day - LocalDBInstance.ApptBitmap_StartDate, a) = Slot_DefaultAccordingToTemplate
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SlotsClear", Err
End Sub

'EHT=Standard
Sub DB_SlotsFill(LocalDBInstance As EJTSClientsDB, Day&, TimeSlot&, NumSlots&, NewApptIndex&)
On Error GoTo ERR_HANDLER

Dim a&, ub&
If Not DB_DayWithinBitmapRange(LocalDBInstance, Day) Then Exit Sub
ub = TimeSlot + NumSlots - 1
If ub > Appointment_NumSlotsUB Then ub = Appointment_NumSlotsUB
For a = TimeSlot To ub
    LocalDBInstance.ApptBitmap(Day - LocalDBInstance.ApptBitmap_StartDate, a) = NewApptIndex
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SlotsFill", Err
End Sub

'EHT=Standard
Sub DB_SlotFill(LocalDBInstance As EJTSClientsDB, Day&, TimeSlot&, NewApptIndex&)
On Error GoTo ERR_HANDLER

If Not DB_DayWithinBitmapRange(LocalDBInstance, Day) Then Exit Sub
LocalDBInstance.ApptBitmap(Day - LocalDBInstance.ApptBitmap_StartDate, TimeSlot) = NewApptIndex

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SlotFill", Err
End Sub

'EHT=Standard
Function DB_SlotsIsAvail(LocalDBInstance As EJTSClientsDB, Day&, TimeSlot&, NumSlots&, IgnoreApptID&) As Boolean
On Error GoTo ERR_HANDLER

Dim a&, ub&, i&
If Not DB_DayWithinBitmapRange(LocalDBInstance, Day) Then Exit Function
ub = TimeSlot + NumSlots - 1
If ub > Appointment_NumSlotsUB Then Exit Function    'Cannot create an appointment that extends outside of the slot range
For a = TimeSlot To ub
    i = LocalDBInstance.ApptBitmap(Day - LocalDBInstance.ApptBitmap_StartDate, a)
    If i >= 0 Then
        If IgnoreApptID < 0 Then
            Exit Function
        Else
            If LocalDBInstance.Appointments(i).ID <> IgnoreApptID Then Exit Function
        End If
    End If
Next a
DB_SlotsIsAvail = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_SlotsIsAvail", Err
End Function

'EHT=Standard
Function DB_FindNextAvailableSlot(LocalDBInstance As EJTSClientsDB, ByVal startdate As Long, NumSlots&, retDay&, retSlot&) As Boolean
On Error GoTo ERR_HANDLER

Dim cd As Long, a&, b&
For cd = startdate - LocalDBInstance.ApptBitmap_StartDate To LocalDBInstance.ApptBitmap_Count - 1
    b = 0
    For a = 0 To Appointment_NumSlotsUB
        If LocalDBInstance.ApptBitmap(cd, a) Then
            b = 0
        Else
            b = b + 1
            If b = NumSlots Then
                retDay = LocalDBInstance.ApptBitmap_StartDate + cd
                retSlot = a - NumSlots + 1
                DB_FindNextAvailableSlot = True
                Exit Function
            End If
        End If
    Next a
Next cd

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_FindNextAvailableSlot", Err
End Function

'EHT=None
Function DB_GetTimeSlotTime(LocalDBInstance As EJTSClientsDB, ts&) As String
DB_GetTimeSlotTime = Format$( _
        CDate((DB_GetSetting(LocalDBInstance, SETTING_FIRSTSLOTTIME) / 24) + _
        (ts * (DB_GetSetting(LocalDBInstance, SETTING_SLOTLENGTH) / 1440))), _
    "h:mm AM/PM")
End Function

'EHT=Standard
Function DB_FormatApptClientList$(LocalDBInstance As EJTSClientsDB, appt As Appointment)
On Error GoTo ERR_HANDLER

Dim a&, cindex&
DB_FormatApptClientList$ = "ApptID#" & appt.ID & "["
For a = 0 To appt.ClientID_Count - 1
    cindex = DB_FindClientIndex(LocalDBInstance, appt.ClientIDs(a))
    If a > 0 Then DB_FormatApptClientList$ = DB_FormatApptClientList$ & " + "
    DB_FormatApptClientList$ = DB_FormatApptClientList$ & LocalDBInstance.Clients(cindex).c.Person1.Last
Next a
DB_FormatApptClientList$ = DB_FormatApptClientList$ & "]"

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_FormatApptClientList", Err
End Function

'EHT=Standard
Function DB_FormatMinutesForSchedule$(LocalDBInstance As EJTSClientsDB, cindex&, primary As Boolean)
On Error GoTo ERR_HANDLER

With LocalDBInstance.Clients(cindex).c
    If Not primary Then DB_FormatMinutesForSchedule$ = "+"
    If Flag_IsSet(.Flags, NewClient) Then
        DB_FormatMinutesForSchedule$ = DB_FormatMinutesForSchedule$ & "N"
    ElseIf Flag_IsSet(.LastYear_Flags, NoNeedToFile) Then
        DB_FormatMinutesForSchedule$ = DB_FormatMinutesForSchedule$ & "NF"
    Else
        DB_FormatMinutesForSchedule$ = DB_FormatMinutesForSchedule$ & FieldToString(.LastYear_MinutesToComplete, mNumberOrNULL)
    End If
End With

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_FormatMinutesForSchedule", Err
End Function

'EHT=Standard
Function DB_GenerateLogfileName$(ByVal f$)
On Error GoTo ERR_HANDLER

Dim p&, p2&
p = InStrRev(f$, ".")
If p > 0 Then
    p2 = InStrRev(f$, "\")
    If p2 < p Then
        f$ = Left$(f$, p - 1)
    End If
End If
DB_GenerateLogfileName$ = f$ & ".log"

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DB_GenerateLogfileName", Err
End Function

