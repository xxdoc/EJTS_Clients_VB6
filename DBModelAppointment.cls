VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelAppointment"

'Back reference to database
Public AttachedDatabase As CDatabase

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Appointment





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, PK1 As Long, PK2 As Long)
Set AttachedDatabase = vDB
'Primary key...
' = PK1                     'PK1, not used
ID = PK2                    'PK2
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set AttachedDatabase = vDB
If schemaversion >= 1 Then
    fh.ReadLong             'PK1, not used
    ID = fh.ReadLong        'PK2
    Get #fh.FileNum, , m    'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong 0              'PK1, not used
fh.WriteLong ID             'PK2
Put #fh.FileNum, , m        'Attr
End Sub

'EHT=None
Public Sub QuickSetData(vDay As Date, vTimeSlot As Long, vNumSlots As Long)
m.Day = vDay
m.TimeSlot = vTimeSlot
m.NumSlots = vNumSlots
End Sub





'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function Clients() As DBModelClient()
Clients = AttachedDatabase.ClientsByAppointment(ID)
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
