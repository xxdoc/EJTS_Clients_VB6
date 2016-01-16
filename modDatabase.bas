Attribute VB_Name = "modDatabase"
Option Explicit

'Client
Public Enum enumMailingListMode
    mlmAuto                         'Default
    mlmNoOrganizer
    mlmEmailOrganizer
    mlmHardCopyOrganizer
End Enum
Public Type typePersonData
    Last As String
    Middle As String                'May be blank
    First As String
    Nickname As String              'May be blank
    Phone As String                 'May be blank
    Email As String                 'May be blank
    DateOfBirth As Long             'May be NullLong
    DateOfDeath As Long             'May be NullLong
End Type
Public Type typeCoreData_Client
    MailingListMode As Long         'enumMailingListMode
    MailingAddress_Street As String
    MailingAddress_City As String
    MailingAddress_State As String  'Two-letter abbreviation
    MailingAddress_ZipCode As String
    HomePhone As String             'May be blank
    NumApptSlots As Long            'Must be at least 1
    ReminderCallAlways As Boolean
    IPTE As Boolean                 'True means it's an Inc/Ptnr/Trust/Estate
    Notes As String                 'May be blank
    OldestYearFiled As Long         'May be NullLong
    NewestYearFiled As Long         'May be NullLong
    PersonCount As Long             'Must be at least 1
    Persons() As typePersonData     'At least one person is required
End Type

'ApptCliLink
Public Enum enumReminderCallStatus
    rcsNoneRequired                 'Default
    rcsNeedToCall
    rcsCalled
End Enum
Public Type typeCoreData_ApptCliLink
    ReminderCallStatus As Long
    NoShow As Boolean
    PrimaryClient As Boolean        'Default=True
End Type

'Appointment
Public Type typeCoreData_Appointment
    Day As Long
    TimeSlot As Long
    CustomTimeSet As Boolean
    CustomTime As Date              'Ignored if CustomTimeSet=False
    NumSlots As Long
    DidntHappen As Boolean
    Notes As String                 'May be blank
End Type

'TaxReturn
Public Enum enumReturnStatus
    rsNotStarted                    'Default
    rsIncomplete
    rsComplete
    rsNoNeedToFile
End Enum
Public Enum enumInboxType
    itAppointment                   'Default
    itDroppedOff
    itMailedIn
End Enum
Public Type typeCoreData_TaxReturn
    Status As Long
    InboxType As Long
    FiledExtension As Boolean
    EFiled As Boolean               'Default=True
    CompletionDate As Long          'May be NullLong
    MinutesToComplete As Long       'May be NullLong
    Fee As Long
    FeeOwed As Long
    ReleasedBeforePayment As Boolean
    StateList As String             'Comma-separated list; blank = no state return
    Results_AGI As Long             'May be NullLong
    Results_Fed As Long             'May be NullLong
    Results_State As Long           'May be NullLong
End Type

'BookkeepingJob
Public Type typeBookkeepingMonth
    CompletionDate As Long          'May be NullLong
    Fee As Long
    FeeOwed As Long
End Type
Public Type typeCoreData_BookkeepingJob
    ClientID As Long                'May be NullLong
    ClientName As String            'May be blank; ignored if ClientID <> NullLong
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
    ClientID As Long                'May be NullLong
    ClientName As String            'May be blank; ignored if ClientID <> NullLong
    Description As String           'May be blank
    CompletionDate As Date
    Fee As Long
    FeeOwed As Long
End Type
