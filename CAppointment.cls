VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CAppointment"

'Back reference to database
Private mDatabase As CDatabase

'Primary key
Public ID As Long

'Attributes
Public Day As Date
Public TimeSlot As Long
Public TimeOverride As Date
Public NumSlots As Long
Public DidntHappen As Boolean
Public Notes As String

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long, vDay As Date, vTimeSlot As Long, vNumSlots As Long)
Set mDatabase = vDB
ID = vID
Day = vDay
TimeSlot = vTimeSlot
NumSlots = vNumSlots
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ID
fh.WriteLong TimeSlot
fh.WriteLong NumSlots
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
ID = fh.ReadLong
TimeSlot = fh.ReadLong
NumSlots = fh.ReadLong
End Sub

'EHT=None
Public Function Clients() As CClient()
Clients = mDatabase.ClientsByAppointment(ID)
End Function
