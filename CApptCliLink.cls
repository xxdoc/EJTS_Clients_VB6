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

'Primary key (for efficiency, make public and do not put within the core data type)
Public ClientID As Long, AppointmentID As Long

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_ApptCliLink





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vAppointmentID As Long)
Set mDatabase = vDB
'Primary key...
ClientID = vClientID: AppointmentID = vAppointmentID
'Defaults...
m.PrimaryClient = True
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
