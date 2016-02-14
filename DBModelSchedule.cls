VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelSchedule"

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public CalendarYear As Long         'PK1 = Must be specified up front
'(PK2 not used)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Schedule





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew(vCalendarYear As Long)
CalendarYear = vCalendarYear
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 1 Then
    CalendarYear = PK(1)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong CalendarYear
For a = 1 To PKCount - 1
    fh.WriteLong 0
Next a
Put #fh.FileNum, , m
End Sub

'EHT=None
Public Sub TEMP_InitWithData(vDB As CDatabase, PK() As Long)
Set AttachedDatabase = vDB
CalendarYear = PK(1)
End Sub
