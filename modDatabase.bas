Attribute VB_Name = "modDatabase"
Option Explicit

'#################################################################################
'Current database structures
'#################################################################################

'Client
Public Enum enumWhichPhoneIsBest
    bcHomePhone                         'Default
    bcPerson1CellPhone
    bcPerson2CellPhone
End Enum
Public Enum enumMailingListMode
    mlmAuto                             'Default
    mlmNoOrganizer
    mlmEmailOrganizer
    mlmHardCopyOrganizer
End Enum
Public Type typePersonData
    First As String                     'May be blank if Last has data
    Nickname As String                  'May be blank
    Middle As String                    'May be blank
    Last As String                      'May be blank if First has data
    CellPhone As String                 'May be blank
    Email As String                     'May be blank
    DateOfBirth As Long                 'May be NullLong
    DateOfDeath As Long                 'May be NullLong
End Type
Public Type typeCoreData_Client
    Unused As Boolean                   'True means this data is just for record keeping, but is no longer an active client, possibly due to being replaced with a new record
    IPTE As Boolean                     'True means it's an Inc/Ptnr/Trust/Estate
    PersonCount As Long                 'Must be at least 1
    Persons() As typePersonData         'At least one person is required
    MailingAddress_Street As String     'May be blank only if City,State,Zip are also blank
    MailingAddress_City As String       'May be blank only if Street,State,Zip are also blank
    MailingAddress_State As String      'May be blank only if Street,City,Zip are also blank; two-letter abbreviation
    MailingAddress_ZipCode As String    'May be blank only if Street,City,State are also blank
    HomePhone As String                 'May be blank
    WhichPhoneIsBest As Long            'enumWhichPhoneIsBest
    NumApptSlots As Long                'Must be at least 1
    ReminderCallAlways As Boolean
    OldestYearFiled As Long             'May be NullLong
    NewestYearFiled As Long             'May be NullLong
    MailingListMode As Long             'enumMailingListMode
    Notes1 As String                    'May be blank
    Notes2 As String                    'May be blank
End Type

'TaxReturn
Public Enum enumReturnStatus
    rsNotStarted                        'Default
    rsIncomplete
    rsComplete
    rsNoNeedToFile
End Enum
Public Enum enumInboxType
    itAppointment                       'Default
    itDroppedOff
    itMailedIn
End Enum
Public Type typeCoreData_TaxReturn
    InboxType As Long
    Status As Long
    FiledExtension As Boolean
    EFiled As Boolean                   'Default=True
    CompletionDate As Long              'May be NullLong
    MinutesToComplete As Long           'May be NullLong
    Fee As Long                         'May be NullLong
    FeeOwed As Long                     'May be NullLong
    ReleasedBeforePayment As Boolean
    StateList As String                 'Comma-separated list; blank = no state return
    Results_AGI As Long                 'May be NullLong
    Results_Fed As Long                 'May be NullLong
    Results_State As Long               'May be NullLong
End Type

'Appointment
Public Type typeCoreData_Appointment
    Day As Long
    TimeSlot As Long
    CustomTimeSet As Boolean
    CustomTime As Date                  'Ignored if CustomTimeSet=False
    NumSlots As Long
    DidntHappen As Boolean
    Notes As String                     'May be blank
End Type

'ApptCliLink
Public Enum enumReminderCallStatus
    rcsNoneRequired                     'Default
    rcsNeedToCall
    rcsCalled
End Enum
Public Type typeCoreData_ApptCliLink
    ReminderCallStatus As Long
    NoShow As Boolean
    PrimaryClient As Boolean            'Default=True
End Type

'Schedule
Public Type typeCoreData_Schedule
    ApptBitmap() As Long
    ApptBitmap_StartDate As Long
    ApptBitmap_Count As Long
    Subtitles() As String               'Same LBound and UBound as ApptBitmap
End Type

'BookkeepingJob
Public Type typeBookkeepingMonth
    CompletionDate As Long              'May be NullLong
    Fee As Long
    FeeOwed As Long
End Type
Public Type typeCoreData_BookkeepingJob
    ClientID As Long                    'May be NullLong
    ClientName As String                'May be blank; ignored if ClientID <> NullLong
    Months(11) As typeBookkeepingMonth
End Type

'FavoriteSearch
Public Type typeCoreData_FavoriteSearch
    DisplayName As String
    ResultsDisplayMode As Long
    SearchString As String
End Type

'ExtraCharge
Public Type typeCoreData_ExtraCharge
    ClientID As Long                    'May be NullLong
    ClientName As String                'May be blank; ignored if ClientID <> NullLong
    Description As String               'May be blank
    CompletionDate As Date
    Fee As Long
    FeeOwed As Long
End Type

Public Const PKCount = 2





'#################################################################################
'Old database structures, for migration purposes
'#################################################################################

'<None yet>
