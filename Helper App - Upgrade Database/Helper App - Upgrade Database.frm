VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EJTSClients Upgrade Database"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFiles 
      Height          =   2535
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   960
      Width           =   7215
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtFolder 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "C:\0Kenneth\Dropbox\EJTSClients Source Code\Data Files (- DEV COPY TO MESS WITH -)\"
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ð"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblNew 
      Caption         =   "v4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label lblOld 
      Alignment       =   1  'Right Justify
      Caption         =   "v3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Convert Database:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'How to use this program:
'- Copy Client_DBPortion from modDB to here and name it Client_DBPortion_OLDVERSION
'- Set OldDBVersion to old version number
'- Increment CurDBFileSpecVersion in modDB to new version number
'- Update modDB.DB_Save to write in the new format
'- Update code below to read from OLD format and convert to NEW format
'- Update modDB.DB_Load to read from new format

Private Const DBFolder$ = "C:\0Kenneth\Visual Basic\Programs\EJTSClients\Data Files (- DEV COPY TO MESS WITH -)\"
Private Const OldDBVersion$ = "EJTS-v03"

Private Type Client_DBPortion_OLDVERSION
    ID As Long
    
    'Profile
    NameLast1 As String
    NameLast2 As String
    NameFirst1 As String
    NameFirst2 As String
    DateOfDeath1 As Long
    DateOfDeath2 As Long
    PhoneH As String            'Phone numbers must be in 0000000000 or 0000000000x* format
    PhoneTPW As String
    PhoneSPW As String
    AddressStreet As String
    AddressCity As String
    AddressState As String
    AddressZipCode As String
    EmailAddress1 As String
    EmailAddress2 As String
    Notes As String
    NumApptSlotsToUse As Long
    Flags As Long
    
    'Last year's data
    LastYear_MinutesToComplete As Long
    LastYear_PrepFee As Long
    LastYear_Flags As Long
    
    'Posting data
    CompletionDate As Long
    MinutesToComplete As Long
    OtherState As String
    PrepFee As Long
    MoneyOwed As Long
    ResultAGI As Long
    ResultFederal As Long
    ResultState As Long
    
    'Operation notes
    OpNotes As String
End Type

Private Sub Form_Load()
txtFolder.Text = DBFolder
lblOld.Caption = OldDBVersion
lblNew.Caption = CurDBFileSpecVersion

Dim curfile$, v$, conv As Boolean
Dim fh As CMNMOD_CFileHandler
curfile$ = Dir$(DBFolder & "\*.dat")
Do Until curfile$ = ""
    Set fh = OpenFile(DBFolder & curfile$, mBinary_Input)
        v$ = Space(8)
        Get #fh.FileNum, , v$
    fh.CloseFile
    
    lstFiles.AddItem curfile$ & vbTab & v$ & IIf(v$ = CurDBFileSpecVersion, "     CONVERTED", "")
    lstFiles.Selected(lstFiles.NewIndex) = (v$ <> CurDBFileSpecVersion)
    
    curfile$ = Dir$
Loop
End Sub










Private Sub btnConvert_Click()
Dim fi&, curfile$, v$, a&, b&, c&, d&, filespec$, success As Boolean
Dim TempDBInstance As EJTSClientsDB
Dim emptyDB As EJTSClientsDB
Dim fh As CMNMOD_CFileHandler
Dim ttta() As Appointment
Dim ttte() As ExtraCharge
Dim tttb() As BookkeepingJob
Dim ttts() As SpecialSearch
Dim tttse() As Setting

btnConvert.Enabled = False
DoEvents

For fi = 0 To lstFiles.ListCount - 1
    If lstFiles.Selected(fi) Then
        success = False
        TempDBInstance = emptyDB
        
        lstFiles.ListIndex = fi
        curfile$ = lstFiles.List(fi)
        curfile$ = Left(curfile$, InStr(curfile$, vbTab) - 1)
        TempDBInstance.DB_FullPath = DBFolder & curfile$
        
        '########## Read old database ##########
        Set fh = OpenFile(TempDBInstance.DB_FullPath, mBinary_Input)
            'FILEHEADER
            filespec$ = fh.ReadString(Len(CurDBFileSpecVersion))
            If filespec$ <> OldDBVersion$ Then
                ShowErrorMsg_Custom "DB_Load", "Database file '" & curfile$ & "' is currently in the " & filespec$ & " format. But this utility will currently only convert from the " & OldDBVersion$ & " format."
                GoTo clnup
            End If
            
            'Clients
            TempDBInstance.Clients_Count = fh.ReadLong
            If TempDBInstance.Clients_Count = 0 Then
                Erase TempDBInstance.Clients
            Else
                ReDim TempDBInstance.Clients(TempDBInstance.Clients_Count - 1)
                For a = 0 To TempDBInstance.Clients_Count - 1
                    Dim ooo As Client_DBPortion_OLDVERSION
                    Dim nnn As Client_DBPortion
                    
                    'Read old structure
                    Get #fh.FileNum, , ooo
                    
                    nnn.ID = ooo.ID
                    
                    ConvertToNewPersonStruct ooo.NameLast1, ooo.NameFirst1, ooo.PhoneTPW, ooo.PhoneSPW, ooo.EmailAddress1, ooo.EmailAddress2, ooo.Flags, nnn.Person1, nnn.Person2
                    nnn.PhoneHome = ooo.PhoneH
                    nnn.AddressStreet = ooo.AddressStreet
                    nnn.AddressCity = ooo.AddressCity
                    nnn.AddressState = ooo.AddressState
                    nnn.AddressZipCode = ooo.AddressZipCode
                    nnn.NumApptSlotsToUse = ooo.NumApptSlotsToUse
                    nnn.Flags = ooo.Flags
                    If ooo.Notes Like "*(MLE)*" Then
                        nnn.MailingListStatus = 1
                        nnn.Notes = Trim(Replace(ooo.Notes, "(MLE)", ""))
                    ElseIf ooo.Notes Like "*(MLH)*" Then
                        nnn.MailingListStatus = 2
                        nnn.Notes = Trim(Replace(ooo.Notes, "(MLH)", ""))
                    ElseIf ooo.Notes Like "*(NML)*" Then
                        nnn.MailingListStatus = 3
                        nnn.Notes = Trim(Replace(ooo.Notes, "(NML)", ""))
                    Else
                        nnn.MailingListStatus = 0
                        nnn.Notes = ooo.Notes
                    End If
                    
                    nnn.LastYear_MinutesToComplete = ooo.LastYear_MinutesToComplete
                    nnn.LastYear_PrepFee = ooo.LastYear_PrepFee
                    nnn.LastYear_Flags = ooo.LastYear_Flags
                    nnn.OldestYearFiled = NullLong
                    nnn.NewestYearFiled = NullLong
                    
                    nnn.CompletionDate = ooo.CompletionDate
                    nnn.MinutesToComplete = ooo.MinutesToComplete
                    nnn.OtherState = ooo.OtherState
                    nnn.PrepFee = ooo.PrepFee
                    nnn.MoneyOwed = ooo.MoneyOwed
                    nnn.ResultAGI = ooo.ResultAGI
                    nnn.ResultFederal = ooo.ResultFederal
                    nnn.ResultState = ooo.ResultState
                    
                    nnn.OpNotes = ooo.OpNotes
                    
                    'Put new structure into database
                    TempDBInstance.Clients(a).c = nnn
                Next a
            End If
            
            'Appointments
            TempDBInstance.Appointments_Count = fh.ReadLong
            If TempDBInstance.Appointments_Count = 0 Then
                Erase TempDBInstance.Appointments
            Else
                ReDim ttta(TempDBInstance.Appointments_Count - 1)
                Get #fh.FileNum, , ttta
                TempDBInstance.Appointments = ttta
            End If
            
            'Appointment Bitmap
            'CAUTION: Indexes are stored in memory, but IDs in database
            TempDBInstance.ApptBitmap_StartDate = fh.ReadLong
            TempDBInstance.ApptBitmap_Count = fh.ReadLong
            If TempDBInstance.ApptBitmap_Count = 0 Then
                Erase TempDBInstance.ApptBitmap
            Else
                ReDim TempDBInstance.ApptBitmap(TempDBInstance.ApptBitmap_Count - 1, Appointment_NumSlotsUB)
                For a = 0 To TempDBInstance.ApptBitmap_Count - 1
                    For b = 0 To Appointment_NumSlotsUB
                        'Convert IDs to Indexes
                        c = fh.ReadLong
                        If c >= 0 Then
                            TempDBInstance.ApptBitmap(a, b) = -1
                            For d = 0 To TempDBInstance.Appointments_Count - 1
                                If TempDBInstance.Appointments(d).ID = c Then
                                    TempDBInstance.ApptBitmap(a, b) = d
                                    Exit For
                                End If
                            Next d
                        Else
                            TempDBInstance.ApptBitmap(a, b) = c
                        End If
                    Next b
                Next a
            End If
            
            'Extra charges
            TempDBInstance.ExtraCharges_Count = fh.ReadLong
            If TempDBInstance.ExtraCharges_Count = 0 Then
                Erase TempDBInstance.ExtraCharges
            Else
                ReDim ttte(TempDBInstance.ExtraCharges_Count - 1)
                Get #fh.FileNum, , ttte
                TempDBInstance.ExtraCharges = ttte
            End If
            
            'Bookkeeping
            TempDBInstance.Bookkeeping_Count = fh.ReadLong
            If TempDBInstance.Bookkeeping_Count = 0 Then
                Erase TempDBInstance.Bookkeeping
            Else
                ReDim tttb(TempDBInstance.Bookkeeping_Count - 1)
                Get #fh.FileNum, , tttb
                TempDBInstance.Bookkeeping = tttb
            End If
            
            'Special searches
            TempDBInstance.SpecialSearches_Count = fh.ReadLong
            If TempDBInstance.SpecialSearches_Count = 0 Then
                Erase TempDBInstance.SpecialSearches
            Else
                ReDim ttts(TempDBInstance.SpecialSearches_Count - 1)
                Get #fh.FileNum, , ttts
                TempDBInstance.SpecialSearches = ttts
            End If
            
            'Settings
            TempDBInstance.Settings_Count = fh.ReadLong
            If TempDBInstance.Settings_Count = 0 Then
                Erase TempDBInstance.Settings
            Else
                ReDim tttse(TempDBInstance.Settings_Count - 1)
                Get #fh.FileNum, , tttse
                TempDBInstance.Settings = tttse
            End If
            
            'FILEFOOTER
            If fh.ReadString(Len(CurDBFileSpecVersion)) <> filespec$ Then
                ShowErrorMsg_Custom "DB_Load", "File footer does not match file header. Unable to read complete file."
                GoTo clnup
            End If
            TempDBInstance.Loaded = True
            TempDBInstance.IsWriteable = True
        fh.CloseFile
        
        
        
        '########## Write new database ##########
        TempDBInstance.DB_FullPath = TempDBInstance.DB_FullPath & ".new"
        DB_Save TempDBInstance
        
        
        '########## Clean up temp files ##########
        RenameFile DBFolder & curfile$, DBFolder & curfile$ & ".old", True
        RenameFile DBFolder & curfile$ & ".new", DBFolder & curfile$, False
        TempDBInstance.DB_FullPath = DBFolder & curfile$
        
        '########## Update list ##########
        Set fh = OpenFile(DBFolder & curfile$, mBinary_Input)
            v$ = Space(8)
            Get #fh.FileNum, , v$
        fh.CloseFile
        lstFiles.List(fi) = curfile$ & vbTab & v$ & IIf(v$ = CurDBFileSpecVersion, "     CONVERTED", "")
        lstFiles.Selected(fi) = (v$ <> CurDBFileSpecVersion)
        
        success = True
        
        
        
clnup:
        fh.CloseFile
        If Not success Then
            If MsgBox("Error loading database file" & "'" & curfile$ & "'. Continue to next one anyway?", vbYesNo Or vbDefaultButton2 Or vbCritical) = vbNo Then
                Exit For
            End If
        End If
        
        DoEvents
    End If
Next fi

btnConvert.Enabled = True
btnConvert.SetFocus
End Sub

Private Sub ConvertToNewPersonStruct(l$, f$, ph1$, ph2$, em1$, em2$, fl As ClientFlags, Person1 As PersonStruct, Person2 As PersonStruct)
'Billman-TestA
'Timothy (Tim) E & Kathleen (Kathy) A [Nichelson-TestB]

Dim t$(), a&, ps&, person(1) As PersonStruct

If Flag_IsSet(fl, ClientFlags.IncPtnrTrustEstate) Then
    person(0).Last = l$
    person(0).First = f$
Else
    'Initially set them both
    person(0).Last = l$
    person(1).Last = l$
    t$ = Split(f$, " ")
    ps = 0
    For a = 0 To UBound(t$)
        If t$(a) = "&" Then
            ps = ps + 1
        ElseIf t$(a) Like "(*)" Then
            person(ps).Nickname = Mid(t$(a), 2, Len(t$(a)) - 2)
        ElseIf t$(a) Like "?" Then
            person(ps).Initial = UCase(t$(a))
        ElseIf t$(a) Like "[[]*[]]" Then
            'This will change the appropriate one
            person(ps).Last = Mid(t$(a), 2, Len(t$(a)) - 2)
        Else
            person(ps).First = t$(a)
        End If
    Next a
    'Then see if they both are still the same. In this case, remove the duplicate
    If person(1).Last = person(0).Last Then person(1).Last = ""
End If

With person(0)
    .Phone = ph1$
    .DOD = NullLong
    .Email = em1$
End With
With person(1)
    .Phone = ph2$
    .DOD = NullLong
    .Email = em2$
End With

Person1 = person(0)
Person2 = person(1)
End Sub




