VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelTaxReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelTaxReturn"

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public TaxYear As Long              'PK1 = Must be specified up front
Public ClientID As Long             'PK2 = Foreign key (DBModelClient.ID)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_TaxReturn





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew(vTaxYear As Long, vClientID As Long)
TaxYear = vTaxYear
ClientID = vClientID
m.Fee = NullLong
m.FeeOwed = NullLong
m.CompletionDate = NullLong
m.MinutesToComplete = NullLong
m.Results_AGI = NullLong
m.Results_Fed = NullLong
m.Results_State = NullLong
m.EFiled = True
End Sub

'EHT=None
Public Sub InitFromSampleData(vDB As CDatabase, PK() As Long, vInboxType As enumInboxType, vStatus As enumReturnStatus)
Set AttachedDatabase = vDB
InitNew PK(1), PK(2)
m.InboxType = vInboxType
m.Status = vStatus
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 1 Then
    TaxYear = PK(1)
    ClientID = PK(2)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong TaxYear
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
'<T2015 for C1234>
ToString = "<T" & TaxYear & " for C" & ClientID & ">"
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get Status() As enumReturnStatus
Status = m.Status
End Property

'EHT=None
Public Property Let Status(nv As enumReturnStatus)
m.Status = nv
End Property





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
