VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBModelClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "DBModelClient"

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public ID As Long                   'PK1 = Autoincrement primary key
'(PK2 not used)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Client

'Cached reference to current year tax return
Public mCurrentYearReturn As DBModelTaxReturn


'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew()
m.NumApptSlots = 2
m.OldestYearFiled = NullLong
m.NewestYearFiled = NullLong
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 1 Then
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

'EHT=None
Public Sub TEMP_InitWithData(vDB As CDatabase, PK() As Long, vFirst1 As String, vFirst2 As String, vLast As String, vEmail As String)
Set AttachedDatabase = vDB
InitNew
ID = PK(1)
PersonCount = 2
m.Persons(0).First = vFirst1
m.Persons(0).Last = vLast
m.Persons(0).Email = vEmail
m.Persons(1).First = vFirst2
End Sub





'#################################################################################
'ToString
'#################################################################################

'EHT=None
Public Function ToString(Optional formatcode As Long) As String
Select Case formatcode
Case 0
    '<C1234 Lastname, First1 & First2>
    ToString = "<C" & ID & " " & m.Persons(0).Last & ", " & m.Persons(0).First
    If m.PersonCount > 1 Then ToString = ToString & " & " & m.Persons(1).First
    ToString = ToString & ">"
Case 100
    'Lastname, First1 & First2
    ToString = m.Persons(0).Last & ", " & m.Persons(0).First
    If m.PersonCount > 1 Then ToString = ToString & " & " & m.Persons(1).First
End Select
End Function





'#################################################################################
'Attributes
'#################################################################################

'EHT=None
Public Property Get PersonCount() As Long
PersonCount = m.PersonCount
End Property

'EHT=None
Public Property Let PersonCount(nv As Long)
Dim a&
If nv > m.PersonCount Then
    ReDim Preserve m.Persons(nv - 1)
    For a = m.PersonCount To nv - 1
        'Initialize the person data
        m.Persons(a).DateOfBirth = NullLong
        m.Persons(a).DateOfDeath = NullLong
    Next a
    m.PersonCount = nv
ElseIf nv < m.PersonCount Then
    ReDim Preserve m.Persons(nv - 1)
    m.PersonCount = nv
End If
End Property

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

'EHT=None
Public Property Get CurrentYearReturn() As DBModelTaxReturn
If mCurrentYearReturn Is Nothing Then Set mCurrentYearReturn = AttachedDatabase.FindTaxReturn(FileToOpen_Year, ID)
Set CurrentYearReturn = mCurrentYearReturn
End Property

'EHT=None
Public Sub FillInNewClientData(t As String)
Dim s() As String
s = Split(t, ",")
If UBound(s) >= 0 Then
    If m.PersonCount < 1 Then PersonCount = 1
    m.Persons(0).Last = Trim$(s(0))
    If UBound(s) > 0 Then
        m.Persons(0).First = Trim$(s(1))
    End If
End If
End Sub





'#################################################################################
'Form
'#################################################################################

'EHT=Standard
Public Function PopulateToForm_Client(frm As frmClientEditPost) As Boolean
On Error GoTo ERR_HANDLER

Dim a As Long, offset As Long, en As Boolean

en = True

'Persons
For a = 0 To m.PersonCount - 1
    offset = a * frm.PersonOffset
    With m.Persons(a)
        FieldToTextbox frm.txtField(fncPerson_First + offset), .First, en
        FieldToTextbox frm.txtField(fncPerson_Nickname + offset), .Nickname, en
        FieldToTextbox frm.txtField(fncPerson_Middle + offset), .Middle, en
        FieldToTextbox frm.txtField(fncPerson_Last + offset), .Last, en
        FieldToTextbox frm.txtField(fncPerson_CellPhone + offset), .CellPhone, en
        FieldToTextbox frm.txtField(fncPerson_Email + offset), .Email, en
        FieldToTextbox frm.txtField(fncPerson_DateOfBirth + offset), .DateOfBirth, en
        FieldToTextbox frm.txtField(fncPerson_DateOfDeath + offset), .DateOfDeath, en
    End With
Next a

'Common
FieldToCheckbox frm.chkField(fncUnused), m.Unused, en
FieldToCheckbox frm.chkField(fncIPTE), m.IPTE, en
FieldToTextbox frm.txtField(fncMailingAddress_Street), m.MailingAddress_Street, en
FieldToTextbox frm.txtField(fncMailingAddress_City), m.MailingAddress_City, en
FieldToTextbox frm.txtField(fncMailingAddress_State), m.MailingAddress_State, en
FieldToTextbox frm.txtField(fncMailingAddress_ZipCode), m.MailingAddress_ZipCode, en
FieldToTextbox frm.txtField(fncHomePhone), m.HomePhone, en
FieldToCheckboxGroup frm.chkgroupField, fncWhichPhoneIsBest, frm.PhoneNumberCount, m.WhichPhoneIsBest, en
FieldToTextbox frm.txtField(fncNumApptSlots), m.NumApptSlots, en
FieldToCheckbox frm.chkField(fncReminderCallAlways), m.ReminderCallAlways, en
FieldToTextbox frm.txtField(fncOldestYearFiled), m.OldestYearFiled, en
FieldToTextbox frm.txtField(fncNewestYearFiled), m.NewestYearFiled, en
FieldToChooser frm.ChooserConfig(fncMailingListMode), m.MailingListMode, en
FieldToTextbox frm.txtField(fncNotes1), m.Notes1, en
FieldToTextbox frm.txtField(fncNotes2), m.Notes2, en

PopulateToForm_Client = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateToForm_Client", Err
End Function

'EHT=Standard
Public Function PopulateToForm_TaxReturn(frm As frmClientEditPost, en As Boolean, TaxReturnYear As Long, TRColumn As Long) As Boolean
On Error GoTo ERR_HANDLER

Dim tr As DBModelTaxReturn, pct As PictureBox, found As Boolean
Dim appt As DBModelAppointment

'Attempt to find the tax return
Set tr = AttachedDatabase.FindTaxReturn(TaxReturnYear, ID)
found = Not tr Is Nothing

frm.lblTaxYear(TRColumn).Caption = TaxReturnYear
frm.lblTaxYear(TRColumn).Enabled = found
frm.lblNoTaxReturn(TRColumn).Visible = Not found
Set pct = frm.pctTaxReturn(TRColumn)
If found Then
    'Found it, so populate the data
    If Not tr.PopulateToForm(frm, TRColumn * frm.TaxReturnOffset, en) Then Exit Function
    'Lookup the most recent appointment for each year
    Set appt = AttachedDatabase.FindMostRecentAppointment(ID, TaxReturnYear + 1)
    If appt Is Nothing Then
        FieldToTextbox frm.txtField(fncTaxReturn_AppointmentDate + (TRColumn * frm.TaxReturnOffset)), "N/A", False
    Else
        FieldToTextbox frm.txtField(fncTaxReturn_AppointmentDate + (TRColumn * frm.TaxReturnOffset)), appt.ToString(100), False
    End If
Else
    'Not found, so hide the frame and show a message instead
    With frm.lblNoTaxReturn(TRColumn)
        .Caption = "There is no tax return entered for this year."
        .Move pct.Left + 8, pct.Top + 8, pct.Width - 16, pct.Height - 16
    End With
End If
pct.Visible = found

PopulateToForm_TaxReturn = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateToForm_TaxReturn", Err
End Function

'EHT=Standard
Public Function PopulateFromForm(frm As frmClientEditPost) As Boolean
On Error GoTo ERR_HANDLER

'Create a temporary structure for data, so if there are any errors the original data is left in a partial state
'DO NOT ACCESS THE `m` STRUCTURE ANYWHERE IN THIS FUNCTION EXCEPT ON THE `m = temp` LINE AT THE END!!
Dim temp As typeCoreData_Client

Dim a As Long, b As Long, offset As Long, tr As DBModelTaxReturn
Dim pd(1) As typePersonData, hasname(1) As Boolean, hasotherdata(1) As Boolean, hasanydata(1) As Boolean

'Load the person data initially into local variables
For a = 0 To 1
    With pd(a)
        offset = a * frm.PersonOffset

        'Retrieve values from textboxes
        If Not ValidateTextbox(frm.txtField(fncPerson_First + offset), .First) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Nickname + offset), .Nickname) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Middle + offset), .Middle) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Last + offset), .Last) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_CellPhone + offset), .CellPhone) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_Email + offset), .Email) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_DateOfBirth + offset), .DateOfBirth) Then GoTo VALIDATION_ERROR
        If Not ValidateTextbox(frm.txtField(fncPerson_DateOfDeath + offset), .DateOfDeath) Then GoTo VALIDATION_ERROR
        hasname(a) = Len(.First) > 0 Or Len(.Last) > 0
        hasotherdata(a) = Len(.Nickname) > 0 Or Len(.Middle) > 0 Or Len(.Email) > 0 Or Len(.CellPhone) > 0 Or .DateOfBirth <> NullLong Or .DateOfDeath <> NullLong
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
If Not ValidateCheckbox(frm.chkField(fncUnused), temp.Unused) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncIPTE), temp.IPTE) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_Street), temp.MailingAddress_Street) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_City), temp.MailingAddress_City) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_State), temp.MailingAddress_State) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncMailingAddress_ZipCode), temp.MailingAddress_ZipCode) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncHomePhone), temp.HomePhone) Then GoTo VALIDATION_ERROR
If Not ValidateCheckboxGroup(frm.chkgroupField, fncWhichPhoneIsBest, frm.PhoneNumberCount, temp.WhichPhoneIsBest) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNumApptSlots), temp.NumApptSlots) Then GoTo VALIDATION_ERROR
If Not ValidateCheckbox(frm.chkField(fncReminderCallAlways), temp.ReminderCallAlways) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncOldestYearFiled), temp.OldestYearFiled) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNewestYearFiled), temp.NewestYearFiled) Then GoTo VALIDATION_ERROR
If Not ValidateChooser(frm.ChooserConfig(fncMailingListMode), temp.MailingListMode) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNotes1), temp.Notes1) Then GoTo VALIDATION_ERROR
If Not ValidateTextbox(frm.txtField(fncNotes2), temp.Notes2) Then GoTo VALIDATION_ERROR

'Tax return (only save current year)
offset = 0
For a = 0 To 0
    If frm.pctTaxReturn(a).Visible Then
        Set tr = AttachedDatabase.FindTaxReturn(FileToOpen_Year - a, ID)
        If tr Is Nothing Then Set tr = New DBModelTaxReturn
        If Not tr.PopulateFromForm(frm, offset) Then Exit Function
    End If
    offset = offset + frm.TaxReturnOffset
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
