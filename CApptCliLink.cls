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

Public Enum enumReminderCallStatus
    rcsNoCallNeeded = 0
    rcsNeeded = 10
    rcsCalled = 20
End Enum

'Back reference to database
Private mDatabase As CDatabase

'Primary key
Public ClientID As Long
Public AppointmentID As Long

'Attributes
Public ReminderCallStatus As enumReminderCallStatus
Public NoShow As Boolean
Public PrimaryClient As Boolean

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vAppointmentID As Long, vPrimaryClient As Boolean)
Set mDatabase = vDB
ClientID = vClientID
AppointmentID = vAppointmentID
PrimaryClient = vPrimaryClient
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ClientID
fh.WriteLong AppointmentID
fh.WriteLong CLng(PrimaryClient)
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
ClientID = fh.ReadLong
AppointmentID = fh.ReadLong
PrimaryClient = fh.ReadLong
End Sub

