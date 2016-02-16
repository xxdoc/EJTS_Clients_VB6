VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelFavoriteSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelFavoriteSearch"

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long                   'PK1 = Autoincrement primary key
'(PK2 not used)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_FavoriteSearch





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew()
End Sub

'EHT=None
Public Sub InitFromSampleData(vDB As CDatabase, PK() As Long, vDisplayName As String)
Set AttachedDatabase = vDB
InitNew
ID = PK(1)
m.DisplayName = vDisplayName
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 2 Then
    ID = PK(1)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong ID
For a = 1 To PKCount - 1
    fh.WriteLong 0
Next a
Put #fh.FileNum, , m
End Sub





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString() As String
'<F1234 Completed Returns>
ToString = "<F" & ID & " " & m.DisplayName & ">"
End Function
