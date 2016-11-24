VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkLabMode 
      Alignment       =   1  'Right Justify
      Caption         =   "Lab Mode"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   5655
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtSMSSuffix 
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
      Left            =   2640
      TabIndex        =   12
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox txtCountryCode 
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
      Left            =   2640
      TabIndex        =   10
      Text            =   "+91"
      Top             =   2424
      Width           =   735
   End
   Begin VB.TextBox txtComPort 
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
      Left            =   2640
      TabIndex        =   8
      Text            =   "4"
      Top             =   1848
      Width           =   735
   End
   Begin VB.TextBox txtDefaultConsultant 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1272
      Width           =   2655
   End
   Begin VB.TextBox txtHeaderSpace 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   696
      Width           =   495
   End
   Begin VB.TextBox txtMapDatabase 
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
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label LabEmail 
      Caption         =   "Email"
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
      Left            =   1200
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label labSMSSuffix 
      Alignment       =   1  'Right Justify
      Caption         =   "SMS Suffix"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label LabCountryCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Country Code"
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
      TabIndex        =   11
      Top             =   2496
      Width           =   1695
   End
   Begin VB.Label LabComPort 
      Alignment       =   1  'Right Justify
      Caption         =   "Com Port number for SMS"
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
      Left            =   0
      TabIndex        =   9
      Top             =   1872
      Width           =   2415
   End
   Begin VB.Label LabDefaultConsultant 
      Alignment       =   1  'Right Justify
      Caption         =   "Default Counsaltant"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1248
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Header Space"
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
      Left            =   720
      TabIndex        =   5
      Top             =   624
      Width           =   1575
   End
   Begin VB.Label LabDatabasePath 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Path Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub CmdBrowse_Click()
On Error GoTo trap
Dim StDataPath As String
Dim inLenfullfile As String

With CommonDialog1
.CancelError = False
.Filter = "Database|ClinicRecords.mdb"
.InitDir = StExePath
.ShowOpen
End With
StDataPath = CommonDialog1.FileName


If Not StDataPath = "" Then
inLenfullfile = Len(StDataPath)
txtMapDatabase.Text = Mid$(StDataPath, 1, inLenfullfile - 17)
End If


Exit Sub
trap:
Err.Clear


End Sub



Private Sub Form_Load()
txtMapDatabase.Text = DataLocation
txtHeaderSpace.Text = intHeaderSpace
txtDefaultConsultant.Text = StDefaultConsultant
txtComPort.Text = intComport
txtCountryCode.Text = CountryCode
TxtSMSSuffix.Text = SMSuffix
txtEmail.Text = DefaultEmail
If LabMode = False Then
ChkLabMode.Value = 0
ElseIf LabMode = True Then
ChkLabMode.Value = 1
End If
End Sub

Private Sub OKButton_Click()
On Error GoTo trap
Dim StLogfile As String
Dim CurFile As New Scripting.FileSystemObject
Dim CurStream  As TextStream
Dim StMessage As String
Dim Labstr As String

If ChkLabMode.Value = 0 Then
Labstr = "False"
ElseIf ChkLabMode.Value = 1 Then
Labstr = "True"
End If

StLogfile = Trim(txtMapDatabase.Text) & vbCrLf & _
            Trim(txtHeaderSpace.Text) & vbCrLf & _
            Trim(txtDefaultConsultant.Text) & vbCrLf & _
            Trim(txtComPort.Text) & vbCrLf & _
            Trim(txtCountryCode.Text) & vbCrLf & _
            Trim(TxtSMSSuffix.Text) & vbCrLf & _
            Trim(txtEmail.Text) & vbCrLf & _
            Labstr
            
            
           
            
If CurFile.FileExists(StExePath & "\LogFile.txt") = False Then

CurFile.CreateTextFile (StExePath & "\LogFile.txt")
End If

Set CurStream = CurFile.OpenTextFile(StExePath & "\LogFile.txt", ForWriting, False)

With CurStream
.WriteLine (StLogfile)
.Close
End With



DataLocation = txtMapDatabase.Text
intHeaderSpace = txtHeaderSpace.Text
StDefaultConsultant = txtDefaultConsultant.Text
intComport = txtComPort.Text
CountryCode = txtCountryCode.Text
SMSuffix = TxtSMSSuffix.Text
DefaultEmail = txtEmail.Text
If ChkLabMode.Value = 0 Then
LabMode = False
ElseIf ChkLabMode.Value = 1 Then
LabMode = True
End If



 StMessage = MsgBox("To apply New Settings Close the Application", vbYesNo, "Close Application")
 If StMessage = vbYes Then
 End
 Else
Me.Hide
 End If
Exit Sub

trap:
MsgBox Err.Description
            


End Sub

Private Sub txtComPort_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub txtHeaderSpace_KeyPress(KeyAscii As Integer)

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If


End Sub
