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

'Primary key (for efficiency, make public and do not put within the core data type)
Public ClientID As Long, TaxYear As Long

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_TaxReturn





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vClientID As Long, vTaxYear As Long)
Set mDatabase = vDB
'Primary key...
TaxYear = vTaxYear: ClientID = vClientID
'Defaults...
m.EFiled = True
m.CompletionDate = NullLong
m.MinutesToComplete = NullLong
m.Results_AGI = NullLong
m.Results_Fed = NullLong
m.Results_State = NullLong
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





'#################################################################################
'Form
'#################################################################################

'EHT=None
Public Sub PopulateForm(frm As frmClientEditPost, offset As Long, enablefields As Boolean)
Dim en As Boolean
en = (offset = 0) And enablefields
FieldToCombobox frm.cboField(fncInboxType + offset), m.InboxType, en
FieldToCombobox frm.cboField(fncStatus + offset), m.Status, en
FieldToCheckbox frm.chkField(fncFiledExtension + offset), m.FiledExtension, en
FieldToTextbox frm.txtField(fncCompletionDate + offset), m.CompletionDate, en
FieldToTextbox frm.txtField(fncMinutesToComplete + offset), m.MinutesToComplete, en
FieldToCheckbox frm.chkField(fncEFiled + offset), m.EFiled, en
FieldToTextbox frm.txtField(fncResultAGI + offset), m.Results_AGI, en
FieldToTextbox frm.txtField(fncResultFederal + offset), m.Results_Fed, en
FieldToTextbox frm.txtField(fncResultState + offset), m.Results_State, en
FieldToTextbox frm.txtField(fncStateList + offset), m.StateList, en
FieldToTextbox frm.txtField(fncFee + offset), m.Fee, en
FieldToTextbox frm.txtField(fncFeeOwed + offset), m.FeeOwed, en
FieldToCheckbox frm.chkField(fncReleasedBeforePayment + offset), m.ReleasedBeforePayment, en
End Sub