VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApptCliLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CApptCliLink"

'Back reference to database
Private mDatabase As CDatabase

Public Enum enumReminderCallStatus
    rcsNoneRequired = 0     'Default
    rcsNeedToCall = 10
    rcsCalled = 20
End Enum

'Primary key (for efficiency, make public and do not put within the core data type)
Public ClientID As Long
Public AppointmentID As Long

'Attributes (for efficiency, store within a type object)
Private Type typeCoreData
    ReminderCallStatus As Long
    NoShow As Boolean
    PrimaryClient As Boolean
End Type
Private m As typeCoreData





'#################################################################################
'Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vAppointmentID As Long, vPrimaryClient As Boolean)
Set mDatabase = vDB
ClientID = vClientID: AppointmentID = vAppointmentID        'PK
'Attr...
m.PrimaryClient = vPrimaryClient
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
If schemaversion >= 1 Then
    ClientID = fh.ReadLong: AppointmentID = fh.ReadLong     'PK
    Get #fh.FileNum, , m                                    'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ClientID: fh.WriteLong AppointmentID           'PK
Put #fh.FileNum, , m                                        'Attr
End Sub





'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function Client() As CClient
Set Client = mDatabase.Client(ClientID)
End Function

'EHT=None
Public Function Appointment() As CAppointment
Set Appointment = mDatabase.Appointment(AppointmentID)
End Function
