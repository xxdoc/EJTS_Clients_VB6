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
    mlAuto
    mlNoOrganizer = 10
    mlEmailOrganizer = 20
    mlHardCopyOrganizer = 30
End Enum

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private Type typePersonData
    NameLast As String
    NameMiddle As String
    NameFirst As String
    Phone As String                 'Formatted as phone
    Email As String                 'Formatted as email
    DateOfBirth As Date             '0=Null
    DateOfDeath As Date             '0=Null
End Type
Private Type typeCoreData
    MailingListMode As Long         'enumMailingListMode, but using a Long because enums aren't allowed
    MailingAddress_Street As String
    MailingAddress_City As String
    MailingAddress_State As String
    MailingAddress_ZipCode As String
    HomePhone As String
    NumApptSlots As Long
    Notes As String
    ReminderCallAlways As Boolean
    PersonCount As Long
    Persons() As typePersonData     '1-based array
End Type
Private m As typeCoreData





'#################################################################################
'Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long, vNameFirst As String, vEmail As String)
Set mDatabase = vDB
ID = vID                    'PK
'Attr...
m.PersonCount = 1
ReDim m.Persons(1 To m.PersonCount)
m.Persons(1).NameFirst = vNameFirst
m.Persons(1).Email = vEmail

m.HomePhone = "1234567890"
m.NumApptSlots = 99
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
    ToString = m.Persons(1).NameLast & ", " & m.Persons(1).NameFirst
    If m.PersonCount > 1 Then ToString = ToString & " & " & m.Persons(2).NameFirst
End Select
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get NameLast(personnum As Long) As String
NameLast = m.Persons(personnum).NameLast
End Property

'EHT=None
Public Property Get NameMiddle(personnum As Long) As String
NameMiddle = m.Persons(personnum).NameMiddle
End Property

'EHT=None
Public Property Get NameFirst(personnum As Long) As String
NameFirst = m.Persons(personnum).NameFirst
End Property
