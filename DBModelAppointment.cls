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
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public CalendarYear As Long         'PK1 = Must be specified up front (not part of the unique portion of the primary key; it is used instead for caching purposes)
Public ID As Long                   'PK2 = Autoincrement primary key

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Appointment





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew(vCalendarYear As Long)
CalendarYear = vCalendarYear
End Sub

'EHT=None
Public Sub InitFromSampleData(vDB As CDatabase, PK() As Long, vDay As Date, vTimeSlot As Long, vNumSlots As Long)
Set AttachedDatabase = vDB
InitNew PK(1)
ID = PK(2)
m.Day = vDay
m.TimeSlot = vTimeSlot
m.NumSlots = vNumSlots
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 2 Then
    CalendarYear = PK(1)
    ID = PK(2)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong CalendarYear
fh.WriteLong ID
For a = 1 To PKCount - 2
    fh.WriteLong 0
Next a
Put #fh.FileNum, , m
End Sub





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString(Optional formatcode As Long) As String
Select Case formatcode
Case 0
    '<A1234 5/6/2015 slot 4 length 2>
    ToString = "<A" & ID & " " & CDate(m.Day) & " slot " & m.TimeSlot & " length " & m.NumSlots & ">"
Case 100
    '5/6/2015 slot 4
    ToString = Format(m.Day, "m/d/yy") & " ts" & m.TimeSlot
End Select
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get Day() As Long
Day = m.Day
End Property

'EHT=None
Public Property Get TimeSlot() As Long
TimeSlot = m.TimeSlot
End Property

'EHT=None
Public Property Get DidntHappen() As Boolean
DidntHappen = m.DidntHappen
End Property
