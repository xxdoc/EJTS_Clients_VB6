VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelApptCliLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelApptCliLink"

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public AppointmentID As Long        'PK1 = Foreign key (DBModelAppointment.ID)
Public ClientID As Long             'PK2 = Foreign key (DBModelClient.ID)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_ApptCliLink





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew(vAppointmentID As Long, vClientID As Long)
AppointmentID = vAppointmentID
ClientID = vClientID
m.PrimaryClient = True
End Sub

'EHT=None
Public Sub InitFromSampleData(vDB As CDatabase, PK() As Long, vPrimaryClient As Boolean)
Set AttachedDatabase = vDB
InitNew PK(1), PK(2)
m.PrimaryClient = vPrimaryClient
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 1 Then
    AppointmentID = PK(1)
    ClientID = PK(2)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong AppointmentID
fh.WriteLong ClientID
For a = 1 To PKCount - 2
    fh.WriteLong 0
Next a
Put #fh.FileNum, , m
End Sub





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString() As String
'{{A1234-C1234}}
ToString = "<A" & AppointmentID & "-link-C" & ClientID & ">"
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get PrimaryClient() As Boolean
PrimaryClient = m.PrimaryClient
End Property

'EHT=None
Public Property Let PrimaryClient(ByVal vNewValue As Boolean)
m.PrimaryClient = vNewValue
End Property
