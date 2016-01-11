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

'Back reference to database
Private mDatabase As CDatabase

'Primary key
Public ClientID As Long
Public TaxYear As Long

'Attributes
Public Status As enumReturnStatus
Public SpecialReturn As Boolean
Public InboxType As enumInboxType
Public FiledExtension As Boolean
Public EFiled As Boolean
Public CompletionDate As Date
Public MinutesToComplete As Long
Public PrepFee As Long
Public PrepFeeOwed As Long
Public ReleasedBeforePayment As Boolean
Public StateList As String 'Comma-separated list; blank = no state return
Public Results_AGI As Long
Public Results_Fed As Long
Public Results_State As Long

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vTaxYear As Long, vInboxType As enumInboxType)
Set mDatabase = vDB
TaxYear = vTaxYear
ClientID = vClientID
InboxType = vInboxType
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ClientID
fh.WriteLong TaxYear
fh.WriteLong InboxType
fh.WriteLong Status
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
ClientID = fh.ReadLong
TaxYear = fh.ReadLong
InboxType = fh.ReadLong
Status = fh.ReadLong
End Sub

