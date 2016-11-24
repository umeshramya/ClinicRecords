VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAppointment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appointment"
   ClientHeight    =   7620
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDAppOID 
      Caption         =   "Appointment ID"
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frmmain 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      Begin VB.Frame FrmShow 
         Height          =   1935
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   6735
         Begin VB.CheckBox ChkPending 
            Caption         =   "Pending (Not seen)"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.ComboBox CmbConsultatnt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            TabIndex        =   14
            Top             =   480
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTP 
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd, dd,MMM,yy, hh:mm tt"
            Format          =   20381699
            CurrentDate     =   41484
         End
         Begin VB.OptionButton Optall 
            Caption         =   "All Cases"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptNew 
            Caption         =   "New Cases"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2100
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton OptOld 
            Caption         =   "Old Cases"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   8
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton CmdShowAppointment 
            Caption         =   "Show Appointment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   7
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label labConsultatnt 
            Caption         =   "Consultant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label LabDaateofAppoitment 
            Caption         =   "Date And Time of Appointment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   3495
         End
      End
      Begin VB.CommandButton CmdCreateAppointment 
         Caption         =   "Create Appointment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TxtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox TxtMobile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label labMobile 
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label LabName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid GridAppointment 
      Height          =   3615
      Left            =   480
      TabIndex        =   12
      Top             =   3840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Appo_ID"
         Caption         =   "Appo_ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Mobile"
         Caption         =   "Mobile"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Appo_Date"
         Caption         =   "Appo_Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dddd, dd,MMM,yyyy, HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Confirmed"
         Caption         =   "Confirmed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Canceled"
         Caption         =   "Canceled"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Old_case"
         Caption         =   "Old_case"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Finished"
         Caption         =   "Finished"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Postponed"
         Caption         =   "Postponed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dddd, dd,MMM,yyyy, HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2745.071
         EndProperty
         BeginProperty Column04 
            Alignment       =   3
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Currecordset As ADODB.Recordset
Dim DocCollection As New Collection
Function getDoctorID(DocName As String) As Integer
Dim StSQL As String
Dim DocId As New ADODB.Recordset

StSQL = "SELECT Doctor_details.Doctor_ID " & _
"From Doctor_details " & _
"WHERE (((Doctor_details.Doctor_Name)='" & DocName & "'));"

With DocId
.Open StSQL, CnnMain, adOpenStatic, adLockReadOnly, adCmdText

If .EOF = False Then
.MoveFirst
getDoctorID = .Fields("Doctor_ID").Value
End If
.Close
End With



End Function


Private Sub CmdAppointment_Click()


End Sub

Private Sub CMDAppOID_Click()
Dim APPID As Integer
Dim StSQL As String

APPID = InputBox("Enter Appointement ID")

StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
"Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
"Appointment.Finished, Appointment.Postponed " & _
"From Appointment " & _
"WHERE (((Appointment.Appo_ID)=" & APPID & "));"


With Currecordset
If .State = adStateOpen Then
.Close
End If
.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText

End With

'GridAppointment.Caption = CmbConsultatnt.Text
Set GridAppointment.DataSource = Currecordset
GridAppointment.Refresh
End Sub

Private Sub CmdCreateAppointment_Click()
Dim StSQL As String

StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
"Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
"Appointment.Finished, Appointment.Postponed " & _
"From Appointment " & _
"WHERE (((Appointment.Appo_ID)=-1));"



With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText
.AddNew
.Fields("Mobile").Value = Trim(TxtMobile.Text)
.Fields("Name").Value = Trim(TxtName.Text)
.Fields("Consultant").Value = getDoctorID(CmbConsultatnt.Text)
.Fields("Appo_Date").Value = DTP.Value
.Fields("Confirmed").Value = False
.Fields("Canceled").Value = False
.Fields("Old_case").Value = False
.Fields("Finished").Value = False

.Update
GridAppointment.Caption = CmbConsultatnt.Text
Set GridAppointment.DataSource = Currecordset
GridAppointment.Refresh


End With


End Sub





Private Sub CmdShowAppointment_Click()
Dim StSQL As String
Dim DocId As Integer
Dim APPDate As Date



APPDate = DTP.Day & " , " & DTP.Month & " , " & DTP.Year
DocId = getDoctorID(CmbConsultatnt.Text)


If ChkPending.Value = vbChecked Then
        If Optall.Value = True Then
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#)AND ((Appointment.Finished)=False));"
        
        
        ElseIf OptNew = True Then
        
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#) AND " & _
            "((Appointment.Old_case)=False)AND ((Appointment.Finished)=False));"
        
        
        ElseIf OptOld.Value = True Then
        
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#) AND " & _
            "((Appointment.Old_case)=True) AND ((Appointment.Finished)=False));"
        
        
        End If

ElseIf ChkPending.Value = vbUnchecked Then

        If Optall.Value = True Then
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#)AND ((Appointment.Finished)=true));"
        
        
        ElseIf OptNew = True Then
        
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#) AND " & _
            "((Appointment.Old_case)=False) AND ((Appointment.Finished)=true));"
        
        
        ElseIf OptOld.Value = True Then
        
            StSQL = "SELECT Appointment.Appo_ID, Appointment.Mobile, Appointment.Name, Appointment.Consultant, " & _
            "Appointment.Appo_Date, Appointment.Confirmed, Appointment.Canceled, Appointment.Old_case, " & _
            "Appointment.Finished, Appointment.Postponed, Appointment.User_Name " & _
            "From Appointment " & _
            "WHERE (((Appointment.Consultant)=" & DocId & ") AND " & _
            "((Appointment.Appo_Date)>= #" & APPDate & "#)AND ((Appointment.Appo_Date)< #" & APPDate + 1 & "#) AND " & _
            "((Appointment.Old_case)=True) AND ((Appointment.Finished)=True));"
        
        
        End If

End If






With Currecordset
If .State = adStateOpen Then
.Close
End If
.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenDynamic, adLockOptimistic, adCmdText

GridAppointment.Caption = CmbConsultatnt.Text
Set GridAppointment.DataSource = Currecordset
GridAppointment.Refresh
End With


End Sub

Private Sub Form_Load()

Dim StSQL As String
Dim DocRecordset As New ADODB.Recordset


Set Currecordset = New ADODB.Recordset
DTP.Value = Now()
CmbConsultatnt.Text = StDefaultConsultant



StSQL = "SELECT Doctor_details.Doctor_ID, Doctor_details.Doctor_Name, Doctor_details.Consultant " & _
"From Doctor_details " & _
"WHERE (((Doctor_details.Consultant)=True));"


With DocRecordset
.Open StSQL, CnnMain, adOpenKeyset, adLockReadOnly, adCmdText


    If .EOF = False Then
    .MoveFirst
    
            Do While .EOF = False
            DocCollection.Add DocRecordset.Fields("Doctor_ID").Value
            CmbConsultatnt.AddItem .Fields("Doctor_Name").Value
            .MoveNext
        Loop
    End If
.Close
End With


End Sub

Private Sub Option3_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set DocCollection = Nothing

End Sub

Private Sub GridAppointment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo trap
Dim CurUser As String
Dim StMessage As String


If KeyCode = 13 Then



StMessage = MsgBox("Selected record will be modified", vbOKCancel, "Modify record")
If StMessage = 1 Then
Currecordset.Update
End If
Else
Currecordset.CancelUpdate
End If


Exit Sub
trap:
MsgBox Err.Description


End Sub

