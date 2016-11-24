Attribute VB_Name = "Search_Module"
Dim DataVar As New Collection
Dim DataType As New Collection

Type typeData
Texttype As String
dateType As String
IntergerType  As String
BooleanType As String
End Type



Public Function add_Sql_collection(strClause As Variant, ClauseType As Variant)
Dim CurClause As Variant
Dim CurClauseType As Variant

CurClause = strClause
CurClauseType = strClauseType


DataVar.Add CurClause
DataType.Add CurClauseType


End Function

Public Function GenerateClause(CollDataVar As Collection, CollDataType As Collection) As String
Dim ColCount As Integer
Dim StCluase As String
collcount = CollDataVar.Count


For i = 0 To ColCount
i = i + 1
CollDataVar.Item (i)
CollDataType.Item (i)


If CollDataType.Item = Texttype Then
'code
'StCluase=
ElseIf CollDataType.Item = dateType Then
'code

ElseIf CollDataType.Item = IntergerType Then
'code


ElseIf CollDataType.Item = BooleanType Then
'code

End If


Next

GenerateClause = StCluase

End Function


