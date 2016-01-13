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

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private Type typeCoreData
    Day As Long
    TimeSlot As Long
    TimeOverride As Date    'May be NullLong
    NumSlots As Long
    DidntHappen As Boolean
    Notes As String         'May be blank
End Type
Private m As typeCoreData





'#################################################################################
'Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long, vDay As Date, vTimeSlot As Long, vNumSlots As Long)
Set mDatabase = vDB
ID = vID                    'PK
'Attr...
m.Day = vDay
m.TimeSlot = vTimeSlot
m.NumSlots = vNumSlots
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
Public Function Clients() As CClient()
Clients = mDatabase.ClientsByAppointment(ID)
End Function





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString(Optional formatcode As Long) As String
Select Case formatcode
Case 0
    '5/6/2015, 4, 2
    ToString = CDate(m.Day) & ", " & m.TimeSlot & ", " & m.NumSlots
End Select
End Function
