VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmNewMedicne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Medicne"
   ClientHeight    =   4635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DBMedicneGrid 
      Height          =   2655
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Trade_Name"
         Caption         =   "Trade_Name"
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
         DataField       =   "Compound"
         Caption         =   "Compound"
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
         DataField       =   "Lab_test"
         Caption         =   "Lab_test"
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
      BeginProperty Column04 
         DataField       =   "User_Name"
         Caption         =   "User_Name"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdSearchCompound 
      Caption         =   "Search Compound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdSearchBrand 
      Caption         =   "Search Brand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox ComCompound 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtBrandName 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton AddNew 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label LabCompound 
      Caption         =   "Compound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label LabTradeName 
      Caption         =   "Brand Name"
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
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmNewMedicne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WithEvents Currecordset As ADODB.Recordset
Attribute Currecordset.VB_VarHelpID = -1
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


Private Sub AddNew_Click()
On Error GoTo trap
Dim StTradeName As String
Dim StSQL As String
StTradeName = TxtBrandName.Text



StSQL = "SELECT Medicines_Table.Trade_Name, Medicines_Table.Compound, Medicines_Table.User_Name, Lab_Test " & _
"From Medicines_Table " & _
"WHERE (((Medicines_Table.Trade_Name) = 'nil'));"
With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText
.AddNew
.Fields("Trade_Name").Value = TxtBrandName.Text
.Fields("Compound").Value = ComCompound.Text
.Fields("User_Name").Value = UserName
If LabMode = True Then
.Fields("Lab_Test").Value = True
ElseIf LabMode = False Then
.Fields("Lab_Test").Value = False
End If


.Update



End With

Set DBMedicneGrid.DataSource = Currecordset
Exit Sub
trap:
MsgBox Err.Description

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

Private Sub CmdSearchBrand_Click()
On Error GoTo trap
Dim StTradeName As String
Dim StSQL As String
StTradeName = TxtBrandName.Text

StSQL = "SELECT Medicines_Table.Trade_Name, Medicines_Table.Compound, Medicines_Table.User_Name " & _
"From Medicines_Table " & _
"WHERE (((Medicines_Table.Trade_Name) Like '" & StTradeName & "%'));"
With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText

End With

Set DBMedicneGrid.DataSource = Currecordset
Exit Sub
trap:
MsgBox Err.Description
End Sub

Private Sub CmdSearchCompound_Click()
On Error GoTo trap
Dim StCompound As String
Dim StSQL As String

StCompound = ComCompound.Text

StSQL = "SELECT Medicines_Table.Trade_Name, Medicines_Table.Compound, Medicines_Table.User_name " & _
"From Medicines_Table " & _
"WHERE (((Medicines_Table.Compound) Like '" & StCompound & "%'));"
With Currecordset
If .State = adStateOpen Then
.Close
End If


.CursorLocation = adUseClient
.Open StSQL, CnnMain, adOpenKeyset, adLockOptimistic, adCmdText

End With

Set DBMedicneGrid.DataSource = Currecordset

Exit Sub
trap:
MsgBox Err.Description

End Sub






Private Sub DBMedicneGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo trap
Dim StMessage As String
Dim CurUser As String
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

Private Sub Form_Load()
Set Currecordset = New ADODB.Recordset
If LabMode = True Then
CmdSearchBrand.Caption = " Search Test"
LabTradeName.Caption = "Lab Test"
CmdSearchCompound.Enabled = False
LabCompound.Caption = "Reference Range"
DBMedicneGrid.Columns.Item(1).Caption = "Lab Test"
DBMedicneGrid.Columns.Item(2).Caption = "Reference Range"

End If
Dim strSql As String
Dim CurPopulate As New ADODB.Recordset

If LabMode = True Then

    strSql = "SELECT DISTINCT Medicines_Table.Compound " & _
    "From Medicines_Table where Lab_test= True " & _
    "ORDER BY Medicines_Table.Compound;"

ElseIf LabMode = False Then
    strSql = "SELECT DISTINCT Medicines_Table.Compound " & _
    "From Medicines_Table where Lab_test= False " & _
    "ORDER BY Medicines_Table.Compound;"

End If
 
 
 
With CurPopulate
.Open strSql, CnnMain, adOpenDynamic, adLockReadOnly, adCmdText
If .EOF = False Then
.MoveLast
.MoveFirst
If Not .RecordCount = 0 Then
Do While Not .EOF
        ComCompound.AddItem (.Fields("Compound").Value)
               
        .MoveNext
    Loop
    End If
    End If
    .Close
    
End With



End Sub

Private Sub Label2_Click()

End Sub
