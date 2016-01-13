VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFavoriteSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CFavoriteSearch"

'Back reference to database
Private mDatabase As CDatabase

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private Type typeCoreData
    DisplayName As String
    ResultsDisplayMode As Long
    SearchString As String
End Type
Private m As typeCoreData
