VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaxReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CTaxReturn"

'Back reference to database
Private mDatabase As CDatabase

Public Enum enumReturnStatus
    rsNotStarted = 0
    rsIncomplete = 10
    rsComplete = 20
End Enum
Public Enum enumInboxType
    itAppointment = 0
    itDroppedOff = 10
    itMailedIn = 20
End Enum

'Primary key (for efficiency, make public and do not put within the core data type)
Public ClientID As Long
Public TaxYear As Long

'Attributes (for efficiency, store within a type object)
Private Type typeCoreData
    Status As Long              'enumReturnStatus, but using a Long because enums aren't allowed
    SpecialReturn As Boolean
    InboxType As Long           'enumInboxType, but using a Long because enums aren't allowed
    FiledExtension As Boolean
    EFiled As Boolean
    CompletionDate As Date
    MinutesToComplete As Long
    PrepFee As Long
    PrepFeeOwed As Long
    ReleasedBeforePayment As Boolean
    StateList As String         'Comma-separated list; blank = no state return
    Results_AGI As Long
    Results_Fed As Long
    Results_State As Long
End Type
Private m As typeCoreData





'#################################################################################
'Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vTaxYear As Long, vInboxType As enumInboxType)
Set mDatabase = vDB
TaxYear = vTaxYear: ClientID = vClientID            'PK
'Attr...
m.InboxType = vInboxType
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
If schemaversion >= 1 Then
    ClientID = fh.ReadLong: TaxYear = fh.ReadLong   'PK
    Get #fh.FileNum, , m                            'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ClientID: fh.WriteLong TaxYear         'PK
Put #fh.FileNum, , m                                'Attr
End Sub





'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function Client() As CClient
Set Client = mDatabase.Client(ClientID)
End Function
