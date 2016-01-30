VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CClient"

'Back reference to database
Private mDatabase As CDatabase

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Client





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub Init(vDB As CDatabase, vID As Long)
Set mDatabase = vDB
'Primary key...
ID = vID
'Defaults...
m.PersonCount = 1
ReDim m.Persons(0)
m.NumApptSlots = 2
m.OldestYearFiled = NullLong
m.NewestYearFiled = NullLong
m.Persons(0).DateOfBirth = NullLong
m.Persons(0).DateOfDeath = NullLong
End Sub

'EHT=None
Public Sub QuickSetData(vFirst As String, vEmail As String)
m.Persons(0).First = vFirst
m.Persons(0).Email = vEmail
End Sub

'EHT=None
Public Sub Unserialize(vDB As CDatabase, schemaversion As Long, fh As CMNMOD_CFileHandler)
Set mDatabase = vDB
If schemaversion >= 1 Then
    ID = fh.ReadLong        'PK
    Get #fh.FileNum, , m    'Attr
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
fh.WriteLong ID             'PK
Put #fh.FileNum, , m        'Attr
End Sub





'#################################################################################
'Relational lookups
'#################################################################################

'EHT=None
Public Function Appointments() As CAppointment()
Appointments = mDatabase.AppointmentsByClient(ID)
End Function

'EHT=None
Public Function ApptCliLinks() As CApptCliLink()
ApptCliLinks = mDatabase.ApptCliLinksByClient(ID)
End Function

'EHT=None
Public Function TaxReturns() As CTaxReturn()
TaxReturns = mDatabase.TaxReturnsByClient(ID)
End Function





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString(Optional formatcode As Long) As String
Select Case formatcode
Case 0
    'Lastname, First1 & First2
    ToString = m.Persons(0).Last & ", " & m.Persons(0).First
    If m.PersonCount > 1 Then ToString = ToString & " & " & m.Persons(1).First
End Select
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get Last(personnum As Long) As String
Last = m.Persons(personnum).Last
End Property

'EHT=None
Public Property Get Middle(personnum As Long) As String
Middle = m.Persons(personnum).Middle
End Property

'EHT=None
Public Property Get First(personnum As Long) As String
First = m.Persons(personnum).First
End Property





'#################################################################################
'Form
'#################################################################################

'EHT=Standard
Public Function PopulateToForm(frm As frmClientEditPost) As Boolean
On Error GoTo ERR_HANDLER

Dim a As Long, offset As Long, tr As CTaxReturn

Dim en As Boolean
en = True

'Persons
For a = 0 To m.PersonCount - 1
    offset = a * frmClientEditPost_PersonOffset
    With m.Persons(a)
        FieldToTextbox frm.txtField(fncPerson_First + offset), .First, en
        FieldToTextbox frm.txtField(fncPerson_Nickname + offset), .Nickname, en
        FieldToTextbox frm.txtField(fncPerson_Middle + offset), .Middle, en
        FieldToTextbox frm.txtField(fncPerson_Last + offset), .Last, en
        FieldToTextbox frm.txtField(fncPerson_Email + offset), .Email, en
        FieldToTextbox frm.txtField(fncPerson_DateOfBirth + offset), .DateOfBirth, en
        FieldToTextbox frm.txtField(fncPerson_DateOfDeath + offset), .DateOfDeath, en
        FieldToTextbox frm.txtField(fncPerson_Phone + offset), .Phone, en
    End With
Next a

'Common
FieldToTextbox frm.txtField(fncMailingAddress_Street), m.MailingAddress_Street, en
FieldToTextbox frm.txtField(fncMailingAddress_City), m.MailingAddress_City, en
FieldToTextbox frm.txtField(fncMailingAddress_State), m.MailingAddress_State, en
FieldToTextbox frm.txtField(fncMailingAddress_ZipCode), m.MailingAddress_ZipCode, en
FieldToTextbox frm.txtField(fncNumApptSlots), m.NumApptSlots, en
FieldToCheckbox frm.chkField(fncReminderCallAlways), m.ReminderCallAlways, en
FieldToCombobox frm.cboField(fncMailingListMode), m.MailingListMode, en
FieldToCheckbox frm.chkField(fncIPTE), m.IPTE, en
FieldToTextbox frm.txtField(fncHomePhone), m.HomePhone, en
FieldToTextbox frm.txtField(fncNotes), m.Notes, en
FieldToTextbox frm.txtField(fncOldestYearFiled), m.OldestYearFiled, en
FieldToTextbox frm.txtField(fncNewestYearFiled), m.NewestYearFiled, en

'TaxReturns
offset = 0
For a = FileToOpen_Year To FileToOpen_Year - 1 Step -1
    frm.lblTaxYear(offset).Caption = a
    Set tr = mDatabase.TaxReturn(ID, a)
    If tr Is Nothing Then
        frm.pctTaxReturn(offset).Visible = False
    Else
        tr.PopulateForm frm, offset, en
        frm.pctTaxReturn(offset).Visible = True
    End If
    offset = offset + frmClientEditPost_TaxReturnOffset
Next a

PopulateToForm = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateToForm", Err
End Function

'EHT=Standard
Public Function PopulateFromForm(frm As frmClientEditPost) As Boolean
On Error GoTo ERR_HANDLER

'Create a temporary structure for data, so if there are any errors the original data is left in a partial state
'DO NOT ACCESS THE `m` STRUCTURE ANYWHERE IN THIS FUNCTION EXCEPT ON THE `m = temp` LINE AT THE END!!
Dim temp As typeCoreData_Client

Dim a As Long, b As Long, offset As Long
Dim pd(1) As typePersonData, isempty(1) As Boolean

'Load the person data initially into local variables
For a = 0 To 1
    With pd(a)
        offset = a * frmClientEditPost_PersonOffset
        FieldFromTextbox frm.txtField(fncPerson_First + offset), .First
        FieldFromTextbox frm.txtField(fncPerson_Nickname + offset), .Nickname
        FieldFromTextbox frm.txtField(fncPerson_Middle + offset), .Middle
        FieldFromTextbox frm.txtField(fncPerson_Last + offset), .Last
        FieldFromTextbox frm.txtField(fncPerson_Email + offset), .Email
        FieldFromTextbox frm.txtField(fncPerson_DateOfBirth + offset), .DateOfBirth
        FieldFromTextbox frm.txtField(fncPerson_DateOfDeath + offset), .DateOfDeath
        FieldFromTextbox frm.txtField(fncPerson_Phone + offset), .Phone
        isempty(a) = (.First = "") And (.Nickname = "") And (.Middle = "") And (.Last = "") And (.Email = "") And _
                     (.DateOfBirth = NullLong) And (.DateOfDeath = NullLong) And (.Phone = "")
    End With
Next a

'Now check which ones have data
temp.PersonCount = 0
b = 0
If isempty(0) Then
    If isempty(1) Then
        Err.Raise 1, , "There must be data for at least one person"
    Else
        'There is only data for Person #2 ...
        temp.PersonCount = 1
        '... so shift that data to Person #1
        b = 1
    End If
Else
    If isempty(1) Then
        'There is only data for Person #1
        temp.PersonCount = 1
    Else
        'There is data for both Persons
        temp.PersonCount = 2
    End If
End If
ReDim temp.Persons(temp.PersonCount - 1)
For a = 0 To temp.PersonCount - 1
    temp.Persons(a) = pd(a + b)
Next a

'Common
FieldFromTextbox frm.txtField(fncMailingAddress_Street), temp.MailingAddress_Street
FieldFromTextbox frm.txtField(fncMailingAddress_City), temp.MailingAddress_City
FieldFromTextbox frm.txtField(fncMailingAddress_State), temp.MailingAddress_State
FieldFromTextbox frm.txtField(fncMailingAddress_ZipCode), temp.MailingAddress_ZipCode
FieldFromTextbox frm.txtField(fncNumApptSlots), temp.NumApptSlots
FieldFromCheckbox frm.chkField(fncReminderCallAlways), temp.ReminderCallAlways
FieldFromCombobox frm.cboField(fncMailingListMode), temp.MailingListMode
FieldFromCheckbox frm.chkField(fncIPTE), temp.IPTE
FieldFromTextbox frm.txtField(fncHomePhone), temp.HomePhone
FieldFromTextbox frm.txtField(fncNotes), temp.Notes
FieldFromTextbox frm.txtField(fncOldestYearFiled), temp.OldestYearFiled
FieldFromTextbox frm.txtField(fncNewestYearFiled), temp.NewestYearFiled

'<Tax Return stuff here>

'If we've gotten this far without errors, then save the data for real
m = temp
PopulateFromForm = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateFromForm", Err
End Function