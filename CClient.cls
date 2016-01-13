VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CClient"

'Back reference to database
Private mDatabase As CDatabase

Public Enum enumMailingListMode
    mlmAuto         'Default
    mlmNoOrganizer
    mlmEmailOrganizer
    mlmHardCopyOrganizer
End Enum

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private Type typePersonData
    Last As String
    Middle As String                'May be blank
    First As String
    Nickname As String              'May be blank
    Phone As String                 'May be blank
    Email As String                 'May be blank
    DateOfBirth As Long             'May be NullLong
    DateOfDeath As Long             'May be NullLong
End Type
Private Type typeCoreData
    MailingListMode As Long
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
Private m As typeCoreData





'#################################################################################
'Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long, vFirst As String, vEmail As String)
Set mDatabase = vDB
ID = vID                    'PK
'Attr...
m.PersonCount = 1
ReDim m.Persons(m.PersonCount - 1)
m.Persons(0).First = vFirst
m.Persons(0).Email = vEmail
'Defaults...
m.NumApptSlots = 2
m.OldestYearFiled = NullLong
m.NewestYearFiled = NullLong
m.Persons(0).DateOfBirth = NullLong
m.Persons(0).DateOfDeath = NullLong
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
If schemaversion >= 1 Then
    ID = fh.ReadLong        'PK
    Get #fh.FileNum, , m    'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ID             'PK
Put #fh.FileNum, , m        'Attr
End Sub





'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function Appointments() As CAppointment()
Appointments = mDatabase.AppointmentsByClient(ID)
End Function

'EHT=None
Public Function ApptCliLinks() As CApptCliLink()
ApptCliLinks = mDatabase.ApptCliLinksByClient(ID)
End Function

'EHT=None
Public Function TaxReturns() As CTaxReturn()
TaxReturns = mDatabase.TaxReturnsByClient(ID)
End Function





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString(Optional formatcode As Long) As String
Select Case formatcode
Case 0
    'Lastname, First1 & First2
    ToString = m.Persons(0).Last & ", " & m.Persons(0).First
    If m.PersonCount > 1 Then ToString = ToString & " & " & m.Persons(1).First
End Select
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get Last(personnum As Long) As String
Last = m.Persons(personnum).Last
End Property

'EHT=None
Public Property Get Middle(personnum As Long) As String
Middle = m.Persons(personnum).Middle
End Property

'EHT=None
Public Property Get First(personnum As Long) As String
First = m.Persons(personnum).First
End Property
