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

'EHT=Standard
Public Function PopulateToForm(frm As frmClientEditPost, offset As Long, enablefields As Boolean) As Boolean
On Error GoTo ERR_HANDLER

Dim en As Boolean
en = (offset = 0) And enablefields
FieldToChooser frm.ChooserConfig(fncTaxReturn_InboxType + offset), m.InboxType, en
FieldToChooser frm.ChooserConfig(fncTaxReturn_Status + offset), m.Status, en
FieldToCheckbox frm.chkField(fncTaxReturn_FiledExtension + offset), m.FiledExtension, en
FieldToTextbox frm.txtField(fncTaxReturn_CompletionDate + offset), m.CompletionDate, en
FieldToTextbox frm.txtField(fncTaxReturn_MinutesToComplete + offset), m.MinutesToComplete, en
FieldToCheckbox frm.chkField(fncTaxReturn_EFiled + offset), m.EFiled, en
FieldToTextbox frm.txtField(fncTaxReturn_ResultAGI + offset), m.Results_AGI, en
FieldToTextbox frm.txtField(fncTaxReturn_ResultFederal + offset), m.Results_Fed, en
FieldToTextbox frm.txtField(fncTaxReturn_ResultStatesCombined + offset), m.Results_State, en
FieldToTextbox frm.txtField(fncTaxReturn_StateList + offset), m.StateList, en
FieldToTextbox frm.txtField(fncTaxReturn_FeeTotal + offset), m.Fee, en
FieldToTextbox frm.txtField(fncTaxReturn_FeeOwed + offset), m.FeeOwed, en
FieldToCheckbox frm.chkField(fncTaxReturn_ReleasedBeforePayment + offset), m.ReleasedBeforePayment, en

PopulateToForm = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateToForm", Err
End Function

'EHT=Standard
Public Function PopulateFromForm(frm As frmClientEditPost, offset As Long) As Boolean
On Error GoTo ERR_HANDLER

Dim temp As typeCoreData_TaxReturn

If Not ValidateChooser(frm.ChooserConfig(fncTaxReturn_InboxType + offset), temp.InboxType) Then GoTo VALIDATION_ERROR
If Not ValidateChooser(frm.ChooserConfig(fncTaxReturn_Status + offset), temp.Status) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncTaxReturn_FiledExtension + offset), temp.FiledExtension) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_CompletionDate + offset), temp.CompletionDate) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_MinutesToComplete + offset), temp.MinutesToComplete) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncTaxReturn_EFiled + offset), temp.EFiled) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_ResultAGI + offset), temp.Results_AGI) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_ResultFederal + offset), temp.Results_Fed) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_ResultStatesCombined + offset), temp.Results_State) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_StateList + offset), temp.StateList) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_FeeTotal + offset), temp.Fee) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncTaxReturn_FeeOwed + offset), temp.FeeOwed) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncTaxReturn_ReleasedBeforePayment + offset), temp.ReleasedBeforePayment) Then GoTo VALIDATION_ERROR

'If we've gotten this far without errors, then save the data for real
m = temp
PopulateFromForm = True

Exit Function
VALIDATION_ERROR:
    ShowErrorMsg "Validation error. Please correct the hilighted fields before continuing."

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateFromForm", Err
End Function
