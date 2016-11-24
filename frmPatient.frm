VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatient 
   Caption         =   "Patient"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComReferringPerVisit 
      Height          =   315
      Left            =   13920
      TabIndex        =   54
      Top             =   480
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox ComConsultant 
      Height          =   315
      Left            =   10920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame FrmCasePaper 
      Caption         =   "Case Paper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   8040
      TabIndex        =   23
      Top             =   720
      Width           =   9015
      Begin VB.Frame FrmLabTest 
         Caption         =   "Lab Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   49
         Top             =   3120
         Width           =   8775
         Begin VB.CommandButton CmdAddtest 
            Caption         =   "Add Test"
            Height          =   375
            Left            =   6480
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox ComLabtest 
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
            Left            =   120
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txttestValue 
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
            Left            =   2280
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label LabLabTest 
            Caption         =   "Test Name                     Value "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.Frame FrmDrgs 
         Caption         =   "Drugs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   8775
         Begin VB.ComboBox CmboDrugType 
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
            ItemData        =   "frmPatient.frx":0000
            Left            =   120
            List            =   "frmPatient.frx":0013
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox CmbBrandName 
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
            Left            =   1155
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox CmbOFrequency 
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
            ItemData        =   "frmPatient.frx":0032
            Left            =   3285
            List            =   "frmPatient.frx":0042
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TxtDurration 
            Height          =   375
            Left            =   4560
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label labprescription 
            Caption         =   "Type          Brand Name                   Frequency     Duration"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.CommandButton CmdCreateTemplate 
         Caption         =   "&Create Template"
         Height          =   495
         Left            =   4800
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton CmdLoadTemplate 
         Caption         =   "Load &Template"
         Height          =   495
         Left            =   3480
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8160
         Width           =   1215
      End
      Begin VB.CommandButton CmdSavePatientRecord 
         Caption         =   "&Save to New Record"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   8160
         Width           =   1815
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Case Paper"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   8160
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox RTxtCasePaper 
         Height          =   6615
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1320
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   11668
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmPatient.frx":0062
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrmDisplay 
      Caption         =   "Display Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   7815
      Begin VB.CommandButton CmdToday 
         Caption         =   "To&day Patients"
         Height          =   375
         Left            =   2280
         TabIndex        =   56
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdolder 
         Caption         =   "&Older"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6120
         TabIndex        =   42
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdRecent 
         Caption         =   "&Recent"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4120
         TabIndex        =   41
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton CmdPatientID 
         Caption         =   "Patient &ID"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOldrecords 
         Caption         =   "Show &Records"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton CmdLoadPatinetFile 
         Caption         =   "&Load Patient File"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid GridPatientRecords 
         Height          =   3975
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
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
         Caption         =   "Patient Records"
         ColumnCount     =   2
         BeginProperty Column00 
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
         BeginProperty Column01 
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameNewPatient 
      Caption         =   "New Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   7815
      Begin VB.TextBox TxtAddress 
         Height          =   1095
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtMobile 
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Frame FrmAge 
         Height          =   855
         Left            =   3480
         TabIndex        =   36
         Top             =   2160
         Width           =   2415
         Begin VB.OptionButton OpDays 
            Caption         =   "Days"
            Height          =   195
            Left            =   1560
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton OpMonths 
            Caption         =   "Months"
            Height          =   195
            Left            =   600
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton OpYears 
            Caption         =   "Years"
            Height          =   195
            Left            =   1560
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtAge 
            Height          =   375
            Left            =   600
            TabIndex        =   9
            Top             =   120
            Width           =   855
         End
         Begin VB.Label labAge 
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame LabSearch 
         Height          =   1815
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   7575
         Begin VB.CommandButton CmdOldPatient 
            Caption         =   "S&earch Old Patinet"
            Height          =   615
            Left            =   5040
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ListBox lisSex 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            ItemData        =   "frmPatient.frx":00E4
            Left            =   960
            List            =   "frmPatient.frx":00EE
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtFirstName 
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
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtMiddleName 
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
            Left            =   2340
            TabIndex        =   5
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox TxtLastName 
            Height          =   495
            Left            =   4560
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label LabLastName 
            Caption         =   "Last Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   35
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label LabMiddleName 
            Caption         =   "Middle Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label labFirstName 
            Caption         =   "First Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label LabSex 
            Caption         =   "Sex"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.ComboBox ComRefDoc 
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
         Left            =   1080
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton CmdNewPatient 
         Caption         =   "&New Patient"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   375
         Left            =   6120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd,MMM,yyyy"
         Format          =   20185091
         UpDown          =   -1  'True
         CurrentDate     =   41484
      End
      Begin VB.Label LabAddress 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label LabMobile 
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label LabRefDoc 
         Caption         =   "Refering Doctor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LabDate 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
   End
   Begin VB.Label LabReferringDoctorpervisit 
      Caption         =   "Referring  Per Episode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   55
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label LabConsultant 
      Caption         =   "Consultant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   25
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label LabCurPatient 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurQuaLification As String

Dim StrFirstName As String
Dim strMiddleName As String
Dim strLastName As String
Dim CurDate As Date
Dim StrReferringPerVisit As String
Dim WithEvents Currecordset  As ADODB.Recordset
Attribute Currecordset.VB_VarHelpID = -1
Private Sub LabTest_populate()
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset



strSql = "SELECT Medicines_Table.Compound, Medicines_Table.Trade_Name, Medicines_Table.Lab_test " & _
"From Medicines_Table " & _
"Where (((Medicines_Table.Lab_test) = True)) " & _
"ORDER BY Medicines_Table.Trade_Name;"




With CurPopulate

.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText
If .RecordCount = 0 Then
Exit Sub
End If
 If .EOF = False Then
.MoveLast
.MoveFirst
Do While Not .EOF
        ComLabtest.AddItem (.Fields("Trade_Name"))
       
               
        .MoveNext
    Loop
    
    End If
    .Close
    
End With



End Sub
Function coluntext(stText As String, inColumns As Integer, LeftRightindent As Integer) As String
Dim PrintWidth As Integer
Dim inSpaceWidth As Integer
Dim inColwidth As Integer
Dim I As Integer
Dim CurColumText As String
Dim IntTextwidth As Integer
Dim IntRemainingSpacePerCol As Integer


Printer.Font.Name = RTxtCasePaper.Font.Name
Printer.Font.Size = RTxtCasePaper.Font.Size

PrintWidth = Printer.Width - LeftRightindent
inColwidth = PrintWidth / inColumns

inSpaceWidth = Printer.TextWidth(" ")
IntTextwidth = Printer.TextWidth(stText)
IntRemainingSpacePerCol = (inColwidth - IntTextwidth) / inSpaceWidth

CurColumText = stText

For I = 1 To IntRemainingSpacePerCol
CurColumText = CurColumText & " "
Next
coluntext = CurColumText

End Function

Function SetInpatientIDrecord()

Dim CurPatientId As Integer
Dim strSql As String

CurPatientId = InputBox("Enter Patient ID", "Ptient ID")
strSql = "SELECT Patient_details.Patient_Id, Patient_details.First_Name, " & _
"Patient_details.Middle_Name, Patient_details.Last_name, Patient_details.Date_Of_Birth, " & _
"Patient_details.Sex, Patient_details.Date_Of_Registretion, Patient_details.Referring_Doctor, " & _
"Patient_details.User_Name, Patient_details.Mobile, Patient_details.Address " & _
"From Patient_details " & _
"WHERE (((Patient_details.Patient_Id)=" & CurPatientId & "));"


With Currecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient
.Open strSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText
.MoveFirst
inPatinet = .Fields("Patient_Id").Value
LabCurPatient.Caption = .Fields("First_Name").Value & " " & .Fields("Middle_Name").Value & " " & .Fields("Last_name").Value & ". and Patient ID is " & inPatinet
'LabCurPatient.Caption = Patient_Name(inPatinet).PatientName & ". and Patient ID is " & inPatinet


End With


End Function

Function UserName_validate() As Boolean
Dim OldUser As String

OldUser = Currecordset.Fields("User_Name").Value

If UserName = "Admin" Then
UserName_validate = True
ElseIf UserName = OldUser Then
UserName_validate = True
Else
UserName_validate = False
End If

End Function




Function DocQuilification(DocName As String) As String
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset


strSql = "SELECT Doctor_details.Qualification " & _
"From Doctor_details " & _
"WHERE (((Doctor_details.Doctor_Name)='" & DocName & "'));"



With CurPopulate

.Open strSql, CnnMain, adOpenStatic, adLockReadOnly, adCmdText
If .EOF = False Then
.MoveFirst
DocQuilification = .Fields("Qualification").Value
End If
.Close

End With

End Function


Function selText_casepaper(addText As String)
RTxtCasePaper.SelText = "  " & addText
End Function


Private Sub ComRefDoc_populate()
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset



strSql = "SELECT Doctor_details.Doctor_Name, Doctor_details.Qualification " & _
"From Doctor_details " & _
"Where (((Doctor_details.Referring_Doctor) = True)) " & _
"ORDER BY Doctor_details.Doctor_Name;"


With CurPopulate

.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText
If .RecordCount = 0 Then
Exit Sub
End If
 If .EOF = False Then
.MoveLast
.MoveFirst
Do While Not .EOF
        ComRefDoc.AddItem (.Fields("Doctor_Name"))
        ComReferringPerVisit.AddItem (.Fields("Doctor_Name"))
               
        .MoveNext
    Loop
    
    End If
    .Close
    
End With



End Sub
Private Sub Comconsultant_populate()
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset



strSql = "SELECT Doctor_details.Doctor_Name, Doctor_details.Consultant " & _
"From Doctor_details " & _
"Where(((Doctor_details.Consultant) = True)) " & _
"ORDER BY Doctor_details.Doctor_Name;"


With CurPopulate
.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText

If .RecordCount = 0 Then
Exit Sub
End If
If .EOF = False Then
.MoveLast
.MoveFirst
Do While Not .EOF
        ComConsultant.AddItem (.Fields("Doctor_Name"))
               
        .MoveNext
    Loop
    
    End If
    .Close
    
End With


End Sub


Private Sub CmbrandName_populate()
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset



strSql = "SELECT Medicines_Table.Trade_Name " & _
"From Medicines_Table where Lab_test = False " & _
"ORDER BY Medicines_Table.Trade_Name;"


With CurPopulate
.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText

If .RecordCount = 0 Then
Exit Sub
End If
If Not .EOF = True Then

.MoveLast
.MoveFirst
Do While Not .EOF
        CmbBrandName.AddItem (.Fields("Trade_Name"))
               
        .MoveNext
    Loop
    .Close
    End If
End With




End Sub


Private Sub AddRefDoc_Click()
MsgBox lisSex.Text
End Sub

Private Sub AddDoc_Click()

End Sub



Private Sub CmbBrandName_KeyUp(KeyCode As Integer, Shift As Integer)
Dim curtext As String
curtext = "    " & Ucase_string(CmbBrandName.Text)

If KeyCode = 13 Then
selText_casepaper curtext

CmbOFrequency.SetFocus
End If

End Sub

Private Sub CmboDrugType_KeyUp(KeyCode As Integer, Shift As Integer)
Dim curtext As String
curtext = Ucase_string(CmboDrugType.Text)


If KeyCode = 13 Then
selText_casepaper curtext

CmbBrandName.SetFocus
End If
End Sub

Private Sub CmbOFrequency_KeyUp(KeyCode As Integer, Shift As Integer)
Dim curtext As String
curtext = "    " & Ucase_string(CmbOFrequency.Text)

If KeyCode = 13 Then
selText_casepaper curtext

TxtDurration.SetFocus
End If
End Sub



Private Sub CmdAddtest_Click()
Dim LabText As String, LabValue As String, LabRange As String
Dim Currecord As New ADODB.Recordset
Dim StSql As String



LabText = Ucase_string(ComLabtest.Text)
LabValue = Ucase_string(txttestValue.Text)





StSql = "SELECT Medicines_Table.Compound " & _
"From Medicines_Table " & _
"WHERE (((Medicines_Table.Trade_Name)='" & LabText & "'));"
Currecord.Open StSql, CnnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not Currecord.EOF = True Then
LabRange = Currecord.Fields("Compound").Value
End If

RTxtCasePaper.SelText = coluntext(LabText, 3, 600)
RTxtCasePaper.SelText = coluntext(LabValue, 3, 600)
RTxtCasePaper.SelText = coluntext(LabRange, 3, 600)
RTxtCasePaper.SelText = vbCrLf & vbCrLf



ComLabtest.SetFocus



End Sub

Private Sub CmdCreateTemplate_Click()
On Error GoTo trap
Dim stFileName As String
 With CommonDialog
 .CancelError = True
 .Filter = "Template|*.mud"
 .InitDir = StExePath
   .ShowSave

  RTxtCasePaper.SaveFile CommonDialog.FileTitle
  End With


Exit Sub
trap:
Err.Clear

End Sub

Private Sub CmdDelete_Click()
On Error GoTo trap
Dim StMessage As String
Dim CurUser As String
Dim StChangeRecord As Boolean


CurUser = Currecordset.Fields("User_Name").Value
StChangeRecord = UserName_validate

If StChangeRecord = False Then
MsgBox "You are not allowed modify the record"
Exit Sub
End If





StMessage = MsgBox("Selected record In grid will be Deleted", vbOKCancel, "Deteted")

If StMessage = 1 Then
Currecordset.Delete
End If




Exit Sub
trap:
MsgBox Err.Description


End Sub


Private Sub CmdLoadPatinetFile_Click()
'On Error GoTo trap
With Currecordset
RTxtCasePaper.Text = .Fields("Case_paper").Value
inPatinet = .Fields("Patient_Id").Value
If IsNull(.Fields("Referring_Doctor").Value) Then
StrReferringPerVisit = ""
Else
StrReferringPerVisit = .Fields("Referring_Doctor").Value
End If
End With
CmdSavePatientRecord.Enabled = False


Exit Sub
trap:
MsgBox Err.Description

End Sub

Private Sub CmdLoadTemplate_Click()
On Error GoTo trap
Dim StTemplate As String
Dim StFirsttext As String, StSecondText As String
StFirsttext = RTxtCasePaper.Text


With CommonDialog
.CancelError = True
.Filter = "Template files File|*.mud"
.InitDir = StExePath
.ShowOpen
StTemplate = .FileName
  RTxtCasePaper.LoadFile CommonDialog.FileTitle
  StSecondText = RTxtCasePaper.Text
  End With
  
RTxtCasePaper.Text = StFirsttext & vbCrLf & StSecondText

Exit Sub

trap:
Err.Clear

End Sub


Private Sub CmdNewPatient_Click()
On Error GoTo trap
Dim StSql As String
Dim stMobile As String

If TxtMobile.Text = "" Then
stMobile = "Not Mentioned"
Else
stMobile = Trim(TxtMobile.Text)
End If



StSql = "SELECT Patient_details.Patient_Id, Patient_details.First_Name, Patient_details.Middle_Name, " & _
"Patient_details.Last_name, Patient_details.Sex, Patient_details.Date_Of_Registretion, " & _
"Patient_details.Date_Of_Birth, Patient_details.Referring_Doctor, " & _
"Patient_details.Mobile, Patient_details.Address, Patient_details.User_Name " & _
"From Patient_details " & _
"WHERE (((Patient_details.Patient_Id)<-1));"



GridPatientRecords.Refresh

With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSql, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText
.AddNew

.Fields("First_Name").Value = Trim(txtFirstName.Text)
.Fields("Middle_Name").Value = Trim(txtMiddleName.Text)
.Fields("Last_name").Value = Trim(TxtLastName.Text)

If OpYears.Value = True Then
.Fields("Date_Of_Birth").Value = Today - (365 * txtAge.Text)

ElseIf OpMonths.Value = True Then
.Fields("Date_Of_Birth").Value = Today - (365 * (1 / 12) * txtAge.Text)
ElseIf OpDays.Value = True Then
.Fields("Date_Of_Birth").Value = Today - txtAge.Text
End If


.Fields("sex").Value = lisSex.Text
.Fields("Date_Of_Registretion").Value = DTPDate.Value
.Fields("Referring_Doctor").Value = ComRefDoc.Text
.Fields("Mobile").Value = stMobile
.Fields("Address").Value = txtAddress.Text
.Fields("User_Name").Value = UserName
.Update
txtFirstName.Text = ""
txtMiddleName.Text = ""
TxtLastName.Text = ""
txtAge.Text = ""
ComRefDoc.Text = ""
txtAddress.Text = ""
TxtMobile.Text = ""


 End With

Set GridPatientRecords.DataSource = Currecordset


inPatinet = GridPatientRecords.Columns("Patient_Id").Value
LabCurPatient.Caption = Patient_Name(inPatinet).PatientName & ". and Patient ID is " & inPatinet

RTxtCasePaper.Text = ""
cmdRecent.Enabled = False
cmdolder.Enabled = False

Exit Sub

trap:
MsgBox Err.Description
Currecordset.CancelUpdate


End Sub

Private Sub cmdolder_Click()
On Error GoTo trap
With Currecordset
If .EOF = True Then
.MoveLast
Else
.MoveNext
End If

RTxtCasePaper.Text = .Fields("Case_paper").Value
inPatinet = .Fields("Patient_Id").Value
If IsNull(.Fields("Referring_Doctor").Value) Then
StrReferringPerVisit = ""
Else
StrReferringPerVisit = .Fields("Referring_Doctor").Value
End If
End With


CmdSavePatientRecord.Enabled = False


Exit Sub
trap:
'MsgBox Err.Description

End Sub

Private Sub CmdOldPatient_Click()
On Error GoTo trap
Dim strSql As String
Dim CurSearch As New Search_Class
Dim FirstName As String
Dim MiddleName As String
Dim LastName As String
Dim Sex As String
Dim ClauseString As String


FirstName = Trim(txtFirstName.Text)
MiddleName = Trim(txtMiddleName.Text)
LastName = Trim(TxtLastName.Text)
Sex = lisSex.Text



CurSearch.add_Sql_collection "Patient_details.First_Name", "Like", "'" & FirstName & "%'"
CurSearch.add_Sql_collection "Patient_details.Middle_Name", "Like", "'" & MiddleName & "%'"
CurSearch.add_Sql_collection "Patient_details.Last_name", "Like", "'" & LastName & "%'"
CurSearch.add_Sql_collection "Patient_details.Sex", "Like", "'" & Sex & "%'"


ClauseString = CurSearch.GenerateClause


strSql = "SELECT Patient_details.Patient_Id, Patient_details.First_Name, " & _
"Patient_details.Middle_Name, Patient_details.Last_name, " & _
"Patient_details.Date_Of_Birth, Patient_details.Sex, " & _
"Patient_details.Date_Of_Registretion, Patient_details.Referring_Doctor, " & _
"Patient_details.User_Name, Patient_details.Mobile, Patient_details.Address " & _
" From Patient_details" & _
" Where (" & ClauseString & ") " & _
"ORDER BY Patient_details.First_Name;"


GridPatientRecords.Refresh

With Currecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient

.Open strSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText
Set GridPatientRecords.DataSource = Currecordset

GridPatientRecords.Refresh
cmdRecent.Enabled = False
cmdolder.Enabled = False



End With
CmdNewPatient.Enabled = False

Exit Sub
trap:
MsgBox Err.Description
End Sub







Private Sub cmdOldrecords_Click()
On Error GoTo trap
Dim strSql As String

If LabMode = False Then
strSql = "SELECT Visit_details.Patient_Id, Visit_details.Visit_Id, Visit_details.Consultant, " & _
"Visit_details.Date_of_Visit, Visit_details.Case_paper, Visit_details.User_Name, Referring_Doctor " & _
"From Visit_details " & _
"WHERE (((Visit_details.Patient_Id)=" & inPatinet & ")) " & _
"ORDER BY Visit_details.Date_of_Visit DESC;"
ElseIf LabMode = True Then
strSql = "SELECT Visit_details.Patient_Id, Visit_details.Visit_Id, Visit_details.Consultant, " & _
"Visit_details.Date_of_Visit, Visit_details.Case_paper, Visit_details.User_Name, Referring_Doctor " & _
"From Visit_details " & _
"WHERE (((Visit_details.Patient_Id)=" & inPatinet & ") And  ((Lab_test) = True)) " & _
"ORDER BY Visit_details.Date_of_Visit DESC;"
End If



With Currecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient
.Open strSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText
Set GridPatientRecords.DataSource = Currecordset

GridPatientRecords.Refresh

End With
CmdLoadPatinetFile.Enabled = True

cmdRecent.Enabled = True

cmdolder.Enabled = True


Exit Sub
trap:
MsgBox Err.Description
End Sub

Private Sub CmdPatientID_Click()
On Error GoTo trap
Dim CurPatientId As Integer
Dim strSql As String

CurPatientId = InputBox("Enter Patient ID", "Ptient ID")
strSql = "SELECT Patient_details.Patient_Id, Patient_details.First_Name, " & _
"Patient_details.Middle_Name, Patient_details.Last_name, Patient_details.Date_Of_Birth, " & _
"Patient_details.Sex, Patient_details.Date_Of_Registretion, Patient_details.Referring_Doctor, " & _
"Patient_details.User_Name, Patient_details.Mobile, Patient_details.Address " & _
"From Patient_details " & _
"WHERE (((Patient_details.Patient_Id)=" & CurPatientId & "));"


With Currecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient
.Open strSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText
Set GridPatientRecords.DataSource = Currecordset

GridPatientRecords.Refresh

End With
cmdRecent.Enabled = False


cmdolder.Enabled = False
Exit Sub
trap:
MsgBox Err.Description


End Sub

Private Sub CmdPrint_Click()
On Error GoTo trap
Dim CurCasePaper As String
Dim StrHeaderspace As String
Dim strCurPatient As String
Dim StMessage As String
Dim I As Integer


If inPatinet <= 0 Then
SetInpatientIDrecord
End If

CurCasePaper = RTxtCasePaper.Text
If CurCasePaper = "" Then
Err.Raise 1, "ClinicRecords", "Case paper can not be nothing"
End If

If CmdSavePatientRecord.Enabled = True Then
    StMessage = MsgBox("Shall I save this record before printing", vbOKCancel, "Save?")
    If StMessage = vbOK Then
    CmdSavePatientRecord_Click
    End If
End If



RTxtCasePaper.Text = ""


  strCurPatient = "Name:- " & Patient_Name(inPatinet).PatientName & vbCrLf
  strCurPatient = strCurPatient & "Age:- : " & CurPatientDetails.Age
  
If LabMode = True And StrReferringPerVisit <> "" Then
    strCurPatient = strCurPatient & "  Sex :- " & CurPatientDetails.Sex & vbCrLf & "Referring Doctor :- " & StrReferringPerVisit & vbCrLf & vbCrLf & vbCrLf
Else
    strCurPatient = strCurPatient & "  Sex :- " & CurPatientDetails.Sex & vbCrLf & vbCrLf & vbCrLf
End If
 
For I = 1 To intHeaderSpace
StrHeaderspace = vbCrLf & StrHeaderspace

Next



With RTxtCasePaper

.SelText = StrHeaderspace & vbCrLf
.SelAlignment = 1 ' right alignment
.SelText = Today & vbCrLf
.SelAlignment = 0 ' left alignment
.SelText = "Patient ID :- " & inPatinet & vbCrLf
.SelText = strCurPatient & CurCasePaper & vbCrLf & vbCrLf
.SelAlignment = 1 'right alignment
.SelText = ComConsultant.Text & vbCrLf
.SelAlignment = 1 'right alignment
.SelText = CurQuaLification & vbCrLf
.SelAlignment = 0 'left alignment


.SelPrint Printer.hDC

End With


CmdSavePatientRecord.Enabled = False
cmdRecent.Enabled = False
cmdolder.Enabled = False
StrReferringPerVisit = ""
Exit Sub
trap:
 MsgBox Err.Description


End Sub











Private Sub cmdRecent_Click()
On Error GoTo trap
With Currecordset

If .BOF = True Then
.MoveFirst
Else
.MovePrevious
End If

RTxtCasePaper.Text = .Fields("Case_paper").Value
inPatinet = .Fields("Patient_Id").Value
If IsNull(.Fields("Referring_Doctor").Value) Then
StrReferringPerVisit = ""
Else
StrReferringPerVisit = .Fields("Referring_Doctor").Value
End If
End With




CmdSavePatientRecord.Enabled = False


Exit Sub
trap:
'MsgBox Err.Description

End Sub

Private Sub CmdSavePatientRecord_Click()
On Error GoTo trap
Dim StSql As String
If RTxtCasePaper.Text = "" Then
MsgBox "Nothing to save"
Exit Sub
End If

StSql = "SELECT Visit_details.Patient_Id, Visit_details.Visit_Id, Visit_details.Consultant, " & _
"Visit_details.Date_of_Visit, Visit_details.Case_paper, Visit_details.User_Name, Lab_Test, Referring_Doctor " & _
"From Visit_details " & _
"WHERE (((Visit_details.Visit_Id)<-1));"

If LabMode = True Then
RTxtCasePaper.SelStart = 0
RTxtCasePaper.SelBold = True
RTxtCasePaper.SelRTF = coluntext("Test Name", 3, 600)
RTxtCasePaper.SelRTF = coluntext("Value", 3, 600)
RTxtCasePaper.SelRTF = coluntext("Normal Range", 3, 600)
RTxtCasePaper.SelRTF = vbCrLf & vbCrLf
RTxtCasePaper.SelBold = False
RTxtCasePaper.SelStart = Len(RTxtCasePaper.TextRTF)

End If


GridPatientRecords.Refresh

With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSql, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText
.AddNew
.Fields("Patient_Id").Value = inPatinet
.Fields("Date_Of_visit").Value = Today
.Fields("Consultant").Value = ComConsultant.Text
.Fields("Case_paper").Value = RTxtCasePaper.Text
StrReferringPerVisit = ComReferringPerVisit.Text
If Not StrReferringPerVisit = "" Then
.Fields("Referring_Doctor").Value = StrReferringPerVisit
End If
If LabMode = True Then
    .Fields("Lab_Test").Value = True
ElseIf LabMode = False Then
    .Fields("Lab_Test").Value = False
End If

.Fields("User_Name").Value = UserName
.Update


 End With

Set GridPatientRecords.DataSource = Currecordset
GridPatientRecords.Refresh

CmdSavePatientRecord.Enabled = False
cmdRecent.Enabled = False
cmdolder.Enabled = False
Exit Sub

trap:
MsgBox Err.Description
Currecordset.CancelUpdate


End Sub


Private Sub CmdSMS_Click()
frmSendSMS.Show
End Sub







Private Sub CmdToday_Click()
Dim StSql As String

StSql = "SELECT Distinct Patient_details.Patient_Id, Patient_details.First_Name, Patient_details.Middle_Name, Patient_details.Last_name, " & _
"Patient_details.Date_Of_Birth, Patient_details.Sex, Patient_details.Date_Of_Registretion, Patient_details.Referring_Doctor, " & _
"Patient_details.User_Name " & _
"FROM Patient_details LEFT JOIN Visit_details ON Patient_details.Patient_Id = Visit_details.Patient_Id " & _
"WHERE (((Patient_details.Date_Of_Registretion)=#" & Today & "#) OR ((Visit_details.Date_of_Visit)= #" & Today & "#));"

GridPatientRecords.Refresh

With Currecordset
If .State = adStateOpen Then
.Close
End If

.CursorLocation = adUseClient

.Open StSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdText
Set GridPatientRecords.DataSource = Currecordset

GridPatientRecords.Refresh
cmdRecent.Enabled = False
cmdolder.Enabled = False



End With
CmdNewPatient.Enabled = False

Exit Sub
trap:
MsgBox Err.Description

End Sub

Private Sub ComConsultant_LostFocus()
CurQuaLification = DocQuilification(ComConsultant.Text)
End Sub



Private Sub Command1_Click()

End Sub

Private Sub ComLabtest_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txttestValue.SetFocus
End If

End Sub



Private Sub CurRecordset_RecordsetChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
CmdLoadPatinetFile.Enabled = False
CmdSavePatientRecord.Enabled = True
inPatinet = -1
LabCurPatient.Caption = ""

End Sub









Private Sub Form_Load()
Set Currecordset = New ADODB.Recordset
DTPDate.Value = Today


       'populating comreferring box

Call ComRefDoc_populate

Call Comconsultant_populate
Call CmbrandName_populate
Call LabTest_populate

RTxtCasePaper.SelIndent = 300
RTxtCasePaper.SelRightIndent = 300
CmdSavePatientRecord.Enabled = False

ComConsultant.Text = StDefaultConsultant
CurQuaLification = DocQuilification(StDefaultConsultant)


If LabMode = False Then
FrmLabTest.Visible = False
ElseIf LabMode = True Then
FrmLabTest.Visible = True
FrmLabTest.Left = FrmDrgs.Left
FrmLabTest.Top = FrmDrgs.Top
FrmLabTest.Width = FrmDrgs.Width
FrmLabTest.Height = FrmDrgs.Height
End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
On Error GoTo trap
Dim StMessage As String

If CmdSavePatientRecord.Enabled = True And RTxtCasePaper.Text <> "" Then


StMessage = MsgBox("Should I save the file to new record", vbYesNoCancel, "Save Patient Record")
        If StMessage = vbYes Then
        Call CmdSavePatientRecord_Click
        ElseIf StMessage = vbCancel Then
        Cancel = True
        ElseIf StMessage = vbNo Then
        Cancel = False
        Unload Me
        Currecordset.Close
        End If
ElseIf CmdSavePatientRecord.Enabled = False Then
    Currecordset.Close
    Unload Me
End If

Exit Sub

trap:
MsgBox Err.Description
End Sub

Private Sub GridPatientRecords_Click()
On Error GoTo trap
inPatinet = GridPatientRecords.Columns("Patient_Id").Value
LabCurPatient.Caption = Patient_Name(inPatinet).PatientName & ". and Patient ID is " & inPatinet
With Currecordset
If IsNull(.Fields("Referring_Doctor").Value) Then
StrReferringPerVisit = ""
Else
StrReferringPerVisit = .Fields("Referring_Doctor").Value
End If
End With
CmdSavePatientRecord.Enabled = True




Exit Sub
trap:
MsgBox Err.Description
End Sub






Private Sub GridPatientRecords_DblClick()
On Error GoTo trap
inPatinet = GridPatientRecords.Columns("Patient_Id").Value
'Patient_Name(inPatinet).PatientName & ". and Patient ID is " & inPatinet

Exit Sub
trap:
MsgBox Err.Description

End Sub

Private Sub GridPatientRecords_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo trap
Dim CurUser As String
Dim StMessage As String
Dim StChangeRecord As Boolean


If KeyCode = 13 Then
CurUser = Currecordset.Fields("User_Name").Value
StChangeRecord = UserName_validate


If StChangeRecord = False Then
MsgBox "You are not allowed modify the record"
Exit Sub
End If


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
















Private Sub RTxtCasePaper_Change()
CmdSavePatientRecord.Enabled = True

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)


If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If


End Sub



Private Sub Option2_Click()

End Sub

Private Sub TxtDurration_KeyUp(KeyCode As Integer, Shift As Integer)

Dim curtext As String

curtext = "    " & Ucase_string(TxtDurration.Text) & vbCrLf & vbCrLf & vbCrLf


If KeyCode = 13 Then
selText_casepaper curtext

CmboDrugType.SetFocus
End If
End Sub


Private Sub txtFirstName_Change()
CmdNewPatient.Enabled = True
End Sub

Private Sub txtFirstName_LostFocus()
txtFirstName.Text = Ucase_string(txtFirstName.Text)



End Sub





Private Sub TxtLastName_LostFocus()
TxtLastName.Text = Ucase_string(TxtLastName.Text)

End Sub

Private Sub txtMiddleName_LostFocus()
txtMiddleName.Text = Ucase_string(txtMiddleName.Text)

End Sub




Private Sub txttestValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
CmdAddtest.SetFocus
End If
End Sub

