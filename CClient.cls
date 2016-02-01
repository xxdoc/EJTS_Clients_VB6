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
        frm.lblTaxYear(offset).Enabled = False
        frm.pctTaxReturn(offset).Visible = False
    Else
        If Not tr.PopulateToForm(frm, offset, en) Then Exit Function
        frm.lblTaxYear(offset).Enabled = True
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

Dim a As Long, b As Long, offset As Long, tr As CTaxReturn
Dim pd(1) As typePersonData, hasname(1) As Boolean, hasotherdata(1) As Boolean, hasanydata(1) As Boolean

'Load the person data initially into local variables
For a = 0 To 1
    With pd(a)
        offset = a * frmClientEditPost_PersonOffset

        'Retrieve values from textboxes
        If Not ValidateTextbox(frm.txtField(fncPerson_First + offset), .First) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Nickname + offset), .Nickname) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Middle + offset), .Middle) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Last + offset), .Last) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Email + offset), .Email) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_DateOfBirth + offset), .DateOfBirth) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_DateOfDeath + offset), .DateOfDeath) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Phone + offset), .Phone) Then GoTo VALIDATION_ERROR
        hasname(a) = Len(.First) > 0 Or Len(.Last) > 0
        hasotherdata(a) = Len(.Nickname) > 0 Or Len(.Middle) > 0 Or Len(.Email) > 0 Or Len(.Phone) > 0 Or .DateOfBirth <> NullLong Or .DateOfDeath <> NullLong
        hasanydata(a) = hasname(a) Or hasotherdata(a)

        'Make sure names are filled out if there is other data filled out
        If hasotherdata(a) And (Not hasname(a)) Then
            ShowErrorMsg "First and/or last names for person #" & a & " are missing."
            Exit Function
        End If
    End With
Next a

'Make sure at least one person has a name (otherwise, it would be hard to find this client record later)
If Not (hasname(0) Or hasname(1)) Then
    ShowErrorMsg "A name must be specified for at least one person."
    Exit Function
End If

'Now check which ones have data, and shift if needed
temp.PersonCount = 0
b = 0
If hasanydata(0) Then
    If hasanydata(1) Then
        'There is data for both Persons
        temp.PersonCount = 2
        
        'If second person has same last name as first person, blank it out
        If LCase$(pd(1).Last) = LCase$(pd(0).Last) Then pd(1).Last = ""
    Else
        'There is only data for Person #1
        temp.PersonCount = 1
    End If
Else
    If hasanydata(1) Then
        'There is only data for Person #2, so shift it to Person #1
        temp.PersonCount = 1
        b = 1
    Else
        'This error should never happen, since we're checking for names above
        Err.Raise 1, , "There must be data for at least one person"
    End If
End If
ReDim temp.Persons(temp.PersonCount - 1)
For a = 0 To temp.PersonCount - 1
    temp.Persons(a) = pd(a + b)
Next a

'Common
If Not ValidateTextbox(frm.txtField(fncMailingAddress_Street), temp.MailingAddress_Street) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_City), temp.MailingAddress_City) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_State), temp.MailingAddress_State) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_ZipCode), temp.MailingAddress_ZipCode) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNumApptSlots), temp.NumApptSlots) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncReminderCallAlways), temp.ReminderCallAlways) Then GoTo VALIDATION_ERROR
If Not ValidateCombobox(frm.cboField(fncMailingListMode), temp.MailingListMode) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncIPTE), temp.IPTE) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncHomePhone), temp.HomePhone) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNotes), temp.Notes) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncOldestYearFiled), temp.OldestYearFiled) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNewestYearFiled), temp.NewestYearFiled) Then GoTo VALIDATION_ERROR

'Tax return (only save current year)
offset = 0
For a = FileToOpen_Year To FileToOpen_Year
    If frm.pctTaxReturn(offset).Visible Then
        Set tr = mDatabase.TaxReturn(ID, a)
        If tr Is Nothing Then Set tr = New CTaxReturn
        If Not tr.PopulateFromForm(frm, offset) Then Exit Function
    End If
    offset = offset + frmClientEditPost_TaxReturnOffset
Next a

'If we've gotten this far without errors, then save the data for real
m = temp
PopulateFromForm = True

Exit Function
VALIDATION_ERROR:
    ShowErrorMsg "Validation error. Please correct the hilighted fields before continuing."

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateFromForm", Err
End Function
