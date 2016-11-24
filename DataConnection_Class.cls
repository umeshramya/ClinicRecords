VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataConnection_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents CnnMain As ADODB.Connection
Attribute CnnMain.VB_VarHelpID = -1

Sub ConnectData()
Set CnnMain = New ADODB.Connection
Dim CurFile As New Scripting.FileSystemObject


If CurFile.FileExists("LogFile.txt") = False Then
CurFile.CreateTextFile ("LogFile.txt")
End If
DataLocation = CurFile.OpenTextFile("Logfile.txt").ReadLine
Dim strSql As String
If CnnMain.State = adStateClosed Then
CnnMain = New ADODB.Connection
CnnMain.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataLocation & ";Persist Security Info=False"
CnnMain.Open
End If
End Sub


