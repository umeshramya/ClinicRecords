VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Search_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'code for class Search_Class

Dim Datatable As New Collection
Dim dataOperator As New Collection
Dim InputVariable As New Collection

'add data variable to these collections
Public Function add_Sql_collection(strtable As Variant, strOperator As Variant, strInputVariable As Variant)

Datatable.Add strtable
dataOperator.Add strOperator
InputVariable.Add strInputVariable
End Function



'generate clause string from here
Function GenerateClause() As String

Dim collcount As Integer
Dim StrRunClause As String
collcount = Datatable.Count
For i = 1 To collcount

StrRunClause = StrRunClause & "((" & Datatable.Item(i) & ") " & dataOperator.Item(i) & " " & InputVariable.Item(i) & ")"
If Not i = collcount Then
StrRunClause = StrRunClause & " And "
End If

Next
GenerateClause = StrRunClause
End Function
