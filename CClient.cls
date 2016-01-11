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

Private Type PersonData
    NameLast As String
    NameMiddle As String
    NameFirst As String
    Phone As String                 'Formatted as phone
    Email As String                 'Formatted as email
    DateOfBirth As Date             '0=Null
    DateOfDeath As Date             '0=Null
End Type

'Back reference to database
Private mDatabase As CDatabase

'Primary key
Public ID As Long

'Attributes
Public MailingListMode As MailingListStatus
Public MailingAddress_Street As String
Public MailingAddress_City As String
Public MailingAddress_State As String
Public MailingAddress_ZipCode As String
Public HomePhone As String
Public NumApptSlots As Long
Public Notes As String
Public ReminderCallAlways As Boolean
Private mPersons() As PersonData     '1-based array
Public PersonCount As Long

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long, vNameFirst As String, vEmail As String)
Set mDatabase = vDB
ID = vID
PersonCount = 1
ReDim mPersons(1 To PersonCount)
mPersons(1).NameFirst = vNameFirst
mPersons(1).Email = vEmail
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ID
fh.WriteLong PersonCount
fh.WriteStringPlusLen mPersons(1).NameFirst
fh.WriteStringPlusLen mPersons(1).Email
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
ID = fh.ReadLong
PersonCount = fh.ReadLong
ReDim mPersons(1 To PersonCount)
mPersons(1).NameFirst = fh.ReadStringPlusLen
mPersons(1).Email = fh.ReadStringPlusLen
End Sub

'EHT=None
Public Property Get NameLast(personnum As Long) As String
NameLast = mPersons(personnum).NameLast
End Property

'EHT=None
Public Property Get NameMiddle(personnum As Long) As String
NameMiddle = mPersons(personnum).NameMiddle
End Property

'EHT=None
Public Property Get NameFirst(personnum As Long) As String
NameFirst = mPersons(personnum).NameFirst
End Property

'EHT=None
Public Property Get FullName(formatting As Long) As String
Select Case formatting
Case 0
    FullName = mPersons(1).NameLast & ", " & mPersons(1).NameFirst & " & " & mPersons(2).NameFirst
End Select
End Property

'EHT=None
Public Function Appointments() As CAppointment()
Appointments = mDatabase.AppointmentsByClient(ID)
End Function
