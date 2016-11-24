VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmNewDoctor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Doctor"
   ClientHeight    =   6705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAddNew 
      Caption         =   "New Doctor"
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   1960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame FrmDoctor 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2415
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
         Left            =   3060
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox ComQualification 
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
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   615
         Left            =   3240
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtHomePhone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtOfficePhone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LabMobile 
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
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LabAddres 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LabHomePhone 
         Caption         =   "Home Phone"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LabOfficePhone 
         Caption         =   "Office phone"
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
         Left            =   5040
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label labQualification 
         Caption         =   "Qualification"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DBDoctor 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Doctor_ID"
         Caption         =   "Doctor_ID"
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
         DataField       =   "Doctor_Name"
         Caption         =   "Doctor_Name"
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
         DataField       =   "Referring_Doctor"
         Caption         =   "Referring_Doctor"
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
      BeginProperty Column03 
         DataField       =   "Consultant"
         Caption         =   "Consultant"
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
      BeginProperty Column04 
         DataField       =   "Qualification"
         Caption         =   "Qualification"
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
      BeginProperty Column05 
         DataField       =   "Address"
         Caption         =   "Address"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "Office_Phone"
         Caption         =   "Office_Phone"
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
      BeginProperty Column08 
         DataField       =   "Home_Phone"
         Caption         =   "Home_Phone"
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
      BeginProperty Column09 
         DataField       =   ""
         Caption         =   ""
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            WrapText        =   -1  'True
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            WrapText        =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column03 
            WrapText        =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            WrapText        =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            WrapText        =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column06 
            WrapText        =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column07 
            WrapText        =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column08 
            WrapText        =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ChkConsultant 
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
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   1160
      Width           =   1455
   End
   Begin VB.CheckBox chkReferring 
      Caption         =   "Referring Doctor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Keep this checked allways"
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmNewDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WithEvents CurRecordset As ADODB.Recordset
Attribute CurRecordset.VB_VarHelpID = -1


Function UserName_validate() As Boolean
Dim OldUser As String

OldUser = CurRecordset.Fields("User_Name").Value

If UserName = "Admin" Then
UserName_validate = True
ElseIf UserName = OldUser Then
UserName_validate = True
Else
UserName_validate = False
End If

End Function



Private Sub CmdAddNew_Click()
On Error GoTo trap
Dim Stsql As String
Dim stQualification As String, stMobile As String
If ComQualification.Text = "" Then
stQualification = "Not Mentioned"
Else
stQualification = ComQualification.Text
End If
If TxtMobile.Text = "" Then
stMobile = "Not Mentioned"
Else
stMobile = TxtMobile.Text
End If


Stsql = "SELECT Doctor_details.Doctor_Name, Doctor_details.Mobile, " & _
"Doctor_details.Referring_Doctor, Doctor_details.Consultant, " & _
"Doctor_details.Qualification, Doctor_details.Address, " & _
"Doctor_details.Office_Phone, Doctor_details.Home_Phone, " & _
"Doctor_details.User_Name " & _
"From Doctor_details " & _
"WHERE (((Doctor_details.Doctor_Name)='dr'));"

With CurRecordset
If .State = adStateOpen Then
.Close
End If



.CursorLocation = adUseClient
.Open Stsql, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText
.AddNew

.Fields("User_Name").Value = UserName
.Fields("Doctor_Name").Value = TxtName.Text
.Fields("Mobile").Value = stMobile
.Fields("Referring_Doctor").Value = chkReferring.Value
.Fields("Consultant").Value = ChkConsultant.Value
.Fields("Qualification").Value = stQualification
.Fields("Address").Value = txtAddress.Text
.Fields("Office_Phone").Value = txtOfficePhone.Text
.Fields("home_phone").Value = txtHomePhone.Text

.Update
End With

Set DBDoctor.DataSource = CurRecordset
 Exit Sub
trap:
 MsgBox Err.Description


End Sub

Private Sub CmdDelete_Click()
On Error GoTo trap
Dim StMessage As String

Dim CurUser As String
Dim StChangeRecord As Boolean


CurUser = CurRecordset.Fields("User_Name").Value
StChangeRecord = UserName_validate

If StChangeRecord = False Then
MsgBox "You are not allowed modify the record"
Exit Sub
End If

StMessage = MsgBox("Selected record In grid will be Deleted", vbOKCancel, "Deteted")

If StMessage = 1 Then
CurRecordset.Delete
End If




Exit Sub
trap:
MsgBox Err.Description

End Sub

Private Sub CmdSearch_Click()
On Error GoTo trap
Dim CurSearch As New Search_Class
Dim StDoctor As String
Dim stMobile As String
Dim stQualification As String
Dim StHomePhone As String
Dim StOfficePhone As String
Dim StAddress As String

Dim Stsql As String
Dim StClause As String

StDoctor = Trim(TxtName.Text)
stQualification = Trim(ComQualification.Text)
stMobile = Trim(TxtMobile.Text)
StHomePhone = Trim(txtHomePhone.Text)
StOfficePhone = Trim(txtOfficePhone.Text)
StAddress = Trim(txtAddress.Text)


With CurSearch
If Not StDoctor = "" Then
.add_Sql_collection "Doctor_details.Doctor_Name", "Like", "'%" & StDoctor & "%'"
End If

If Not stQualification = "" Then
.add_Sql_collection "Doctor_details.Qualification", "Like", "'%" & stQualification & "%'"
End If

If Not stMobile = "" Then
.add_Sql_collection "Doctor_details.Mobile", "Like", "'%" & stMobile & "%'"
End If

If Not StHomePhone = "" Then
.add_Sql_collection "Doctor_details.Home_Phone", "Like", "'%" & StHomePhone & "%'"
End If

If Not StOfficePhone = "" Then
.add_Sql_collection "Doctor_details.Office_Phone", "Like", "'%" & StOfficePhone & "%'"
End If

If Not StAddress = "" Then
.add_Sql_collection "Doctor_details.Address", "Like", "'%" & StAddress & "%'"
End If

StClause = .GenerateClause
End With


Stsql = "SELECT Doctor_details.Doctor_ID, Doctor_details.Doctor_Name, Doctor_details.Referring_Doctor, " & _
"Doctor_details.Consultant, Doctor_details.Qualification, Doctor_details.Address, " & _
"Doctor_details.Mobile, Doctor_details.Office_Phone, Doctor_details.Home_Phone, Doctor_details.User_Name " & _
"FROM Doctor_details " & _
"WHERE (" & StClause & ");"


DBDoctor.Refresh

With CurRecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient


.Open Stsql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText

Set DBDoctor.DataSource = CurRecordset

DBDoctor.Refresh

End With


Exit Sub
trap:

MsgBox Err.Description


End Sub



Private Sub DBDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo trap
Dim StMessage As String
Dim CurUser As String
Dim StChangeRecord As Boolean


If KeyCode = 13 Then
CurUser = CurRecordset.Fields("User_Name").Value
StChangeRecord = UserName_validate


If StChangeRecord = False Then
MsgBox "You are not allowed modify the record"
Exit Sub
End If

StMessage = MsgBox("Selected record will be modified", vbOKCancel, "Modify record")
If StMessage = 1 Then
CurRecordset.Update
End If
Else
CurRecordset.CancelUpdate
End If


Exit Sub
trap:
MsgBox Err.Description


End Sub

Private Sub Form_Load()
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset
Set CurRecordset = New ADODB.Recordset
chkReferring.Value = 1

strSql = "SELECT DISTINCT Doctor_details.Qualification " & _
"FROM Doctor_details;"


With CurPopulate

.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText
If .RecordCount = 0 Then
Exit Sub
End If
If .EOF = False Then
.MoveLast
.MoveFirst

Do While Not .EOF
        ComQualification.AddItem (.Fields("Qualification").Value)
    
        .MoveNext
    Loop
    End If
    .Close
    
End With



End Sub

