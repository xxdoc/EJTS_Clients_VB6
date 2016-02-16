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

Private Const errOutsideBitmapRange = "Date and/or timeslot specified are outside the schedule's range."
Private Const errAppointmentMustBeAdded = "Appointment must be added to the database before being added to the schedule."
Private Const SlotCount = 14
Private Const FirstSlotTime = 9 / 24
Private Const SlotDuration = 0.75 / 24

'Back reference to database
Public AttachedDatabase As CDatabase
Public Deleted As Boolean

'Primary key (for efficiency, make public and do not put within the core data type)
Public CalendarYear As Long         'PK1 = Must be specified up front
'(PK2 not used)

'Attributes (for efficiency, store within a type object)
Private m As typeCoreData_Schedule
Private AppointmentIndexes() As Long    'Used only for caching the IDs in m.AppointmentIDs





'#################################################################################
'Initialization & Serialization
'#################################################################################

'EHT=None
Public Sub InitNew(vCalendarYear As Long)
Dim ed As Long
CalendarYear = vCalendarYear
m.StartDate = DateSerial(vCalendarYear, 1, 1)
ed = DateSerial(vCalendarYear, 12, 31)
m.DayCount = ed - m.StartDate + 1
ReDim m.AppointmentIDs(m.DayCount - 1, SlotCount - 1)
ReDim AppointmentIndexes(m.DayCount - 1, SlotCount - 1)
ReDim m.Subtitles(m.DayCount - 1)
End Sub

'EHT=None
Public Sub InitFromSampleData(vDB As CDatabase, PK() As Long)
Set AttachedDatabase = vDB
InitNew PK(1)
End Sub

'EHT=None
Public Sub InitFromFile(fh As CMNMOD_CFileHandler, schemaversion As Long, PK() As Long)
If schemaversion >= 4 Then
    CalendarYear = PK(1)
    Get #fh.FileNum, , m
End If
End Sub

'EHT=None
Public Sub Serialize(fh As CMNMOD_CFileHandler)
Dim a As Long
fh.WriteLong CalendarYear
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
'<Cal2015>
ToString = "<Cal" & CalendarYear & ">"
End Function





'#################################################################################
'Functions
'#################################################################################

'EHT=None
Public Function SlotsAvailable(vDay&, vTimeSlot&, vNumSlots&, vIgnoreApptID&) As Boolean
Dim d&, s&, ub&, aid&
d = vDay - m.StartDate
ub = vTimeSlot + vNumSlots - 1
If ((d < 0) Or (d >= m.DayCount) Or (vTimeSlot < 0) Or (ub > SlotCount)) Then Err.Raise 1, , errOutsideBitmapRange
For s = vTimeSlot To ub
    aid = m.AppointmentIDs(d, s)
    If aid >= 0 Then
        'Positive numbers indicate a real AppointmentID, so this slot is NOT available
        If vIgnoreApptID < 0 Then
            Exit Function
        Else
            'Unless the slot is filled with vIgnoreApptID, in which case we'll pretend it's available
            If aid <> vIgnoreApptID Then Exit Function
        End If
    End If
Next s
SlotsAvailable = True
End Function

'EHT=None
Public Sub FillSlotsWithAppt(vDay&, vTimeSlot&, vNumSlots&, vNewAppt As DBModelAppointment)
Dim d&, s&, ub&, aindex&
d = vDay - m.StartDate
ub = vTimeSlot + vNumSlots - 1
If ((d < 0) Or (d >= m.DayCount) Or (vTimeSlot < 0) Or (ub > SlotCount)) Then Err.Raise 1, , errOutsideBitmapRange
If AttachedDatabase.FindAppointment(vNewAppt.ID, aindex) Is Nothing Then Err.Raise 1, , errAppointmentMustBeAdded
For s = vTimeSlot To ub
    m.AppointmentIDs(d, s) = vNewAppt.ID
    AppointmentIndexes(d, s) = aindex
Next s
End Sub

'EHT=None
Public Sub FillSlotsWithSpecialValue(vDay&, vTimeSlot&, vNumSlots&, vNewValue As enumSlotSpecialValue)
Dim d&, s&, ub&
d = vDay - m.StartDate
ub = vTimeSlot + vNumSlots - 1
If ((d < 0) Or (d >= m.DayCount) Or (vTimeSlot < 0) Or (ub > SlotCount)) Then Err.Raise 1, , errOutsideBitmapRange
For s = vTimeSlot To ub
    m.AppointmentIDs(d, s) = vNewValue
    AppointmentIndexes(d, s) = 0
Next s
End Sub

'EHT=None
Function GetTimeSlotTime(ts&) As String
GetTimeSlotTime = Format$(CDate(FirstSlotTime + (ts * SlotDuration)), "h:mm AM/PM")
End Function
