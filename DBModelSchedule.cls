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

'Primary key (for efficiency, make public and do not put within the core data type)
Public TaxYear As Long

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Schedule





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vTaxYear As Long)
Set AttachedDatabase = vDB
'Primary key...
TaxYear = vTaxYear
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set AttachedDatabase = vDB
If schemaversion >= 1 Then
    TaxYear = fh.ReadLong   'PK
    Get #fh.FileNum, , m    'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong TaxYear        'PK
Put #fh.FileNum, , m        'Attr
End Sub
