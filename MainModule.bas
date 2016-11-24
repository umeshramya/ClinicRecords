Attribute VB_Name = "MainModule"
Public inPatinet As Long
Public LoginSucceeded As Boolean
Public CnnMain As ADODB.Connection
Public UserName As String, StrPassword As String
Public DataLocation As String
Public Today As Date
Public intHeaderSpace As Integer
Public StDefaultConsultant As String
Public intComport As Integer
Public CountryCode  As String
Public SMSuffix As String
Public DefaultEmail As String
Public StExePath As String
Public LabMode As Boolean



Public Type Patient_Matrix
PatientID As Long
Prefix_To_name As String
PatientName As String
Referring_Doctor As String
Age As Variant
Sex As String
Weight As Byte
Height As Byte
End Type


'CurPatientDetails thisn is created for avoiding repeated trips to database _
from Patient_name patient_matrix datatype
Public CurPatientDetails As Patient_Matrix


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Sub Main()
On Error GoTo trap
Dim CurFile As New Scripting.FileSystemObject
Dim CurStream  As TextStream
Dim Currecordset As New ADODB.Recordset
Dim strConn As String
Dim StLogfile As String

Today = Date

LoginSucceeded = False


Set CnnMain = New ADODB.Connection
StExePath = CurFile.GetAbsolutePathName("clinicrecords.exe")
StExePath = App.Path


If CurFile.FileExists(StExePath & "\LogFile.txt") = False Then
CurFile.CreateTextFile (StExePath & "\LogFile.txt")


Set CurStream = CurFile.OpenTextFile(StExePath & "\LogFile.txt", ForWriting, False)


'setting default setteing for lig file
With CurStream
DataLocation = ""
intHeaderSpace = 7
StDefaultConsultant = ""
intComport = 4
CountryCode = "+91"
SMSuffix = ""
DefaultEmail = ""

End With
StLogfile = DataLocation & vbCrLf & _
            intHeaderSpace & vbCrLf & _
            StDefaultConsultant & vbCrLf & _
            intComport & vbCrLf & _
            CountryCode & vbCrLf & _
            SMSuffix & _
            DefaultEmail & _
            LabMode
CurStream.WriteLine (StLogfile)
CurStream.Close

Else
Set CurStream = CurFile.OpenTextFile(StExePath & "\LogFile.txt")

'setting default setteing for lig file
With CurStream
DataLocation = .ReadLine
intHeaderSpace = .ReadLine
StDefaultConsultant = .ReadLine
intComport = .ReadLine
CountryCode = .ReadLine
SMSuffix = .ReadLine
DefaultEmail = .ReadLine
LabMode = .ReadLine

End With

End If




If CnnMain.State = adStateClosed Then

CnnMain = New ADODB.Connection
With CnnMain
.Provider = "Microsoft.Jet.OLEDB.4.0"
   .Properties("Data Source") = DataLocation & "ClinicRecords.mdb"
   .Properties("Jet OLEDB:System database") = DataLocation & "clincrecordsSecurity.mdw"
   .Open UserID:=UserName, Password:=StrPassword
   
End With
   


'CnnMain.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataLocation & "ClinicRecords.mdb" & ";Persist Security Info=False"
'CnnMain.Open
End If

If CnnMain.State = adStateOpen Then
LoginSucceeded = True
End If
Exit Sub

trap:
FrmSetting.Show




End Sub

Function Ucase_string(Curstring As String) As String
Dim t As String
t = Curstring
If t <> "" Then
Mid$(t, 1, 1) = UCase$(Mid$(t, 1, 1))
For I = 1 To Len(t) - 1
If Mid$(t, 1, 2) = Chr$(13) + Chr$(10) Then
Mid$(t, 1 + 2, 1) = UCase$(Mid$(t, I + 2, 1))
End If
If Mid$(t, 1, 1) = "" Then
Mid$(t, I + 1, 1) = UCase$(Mid$(t, I + 1))
End If
Next

Ucase_string = t
End If




End Function

Function Patient_Name(ByVal PatientID As Integer) As Patient_Matrix
Dim nameRecord As New ADODB.Recordset
Dim strSql As String
Dim FirstName As String, MiddleName As String, LastName As String, StrSex As String
Dim inAge As Variant
Dim CurMatrix As Patient_Matrix
Dim BirthDate As Date
Dim strPrefix As String



strSql = "SELECT Patient_details.Patient_Id, Patient_details.First_Name, " & _
"Patient_details.Middle_Name, Patient_details.Last_name, Patient_details.Date_Of_Birth, " & _
"Patient_details.Sex " & _
"From Patient_Details " & _
"WHERE (((Patient_details.Patient_Id)=" & PatientID & "));"



With nameRecord
.Open strSql, CnnMain, adOpenKeyset, adLockPessimistic, adCmdTex

FirstName = .Fields("First_Name").Value
MiddleName = .Fields("Middle_Name").Value
LastName = .Fields("Last_name").Value

StrSex = .Fields("Sex").Value

BirthDate = .Fields("Date_Of_Birth").Value
inAge = Year(Today) - Year(BirthDate) & " Years"

If inAge = "0 Years" Then
If Today - BirthDate < 31 Then
inAge = Today - BirthDate & " Days"
Else
inAge = Month(Today) - Month(BirthDate) & " Months"
End If
End If

'If inAge = "0 Months" Then
'inAge = Day(Today) - Day(BirthDate) & " Days"
'End If
.Close
End With

'prefix to name code
If Year(Today) - Year(BirthDate) < 15 Then
    If StrSex = "Male" Then
    strPrefix = "Master"
    Else
    strPrefix = "Miss"
    End If
Else
    If StrSex = "Male" Then
    strPrefix = "Mr"
    Else
    strPrefix = "Ms"
    End If
End If

With CurPatientDetails
.PatientID = PatientID
.PatientName = strPrefix & " " & Ucase_string(FirstName) & " " & Ucase_string(MiddleName) & " " & Ucase_string(LastName)
.Sex = StrSex
.Age = inAge
.Prefix_To_name = strPrefix
End With

With Patient_Name

.PatientID = PatientID
.PatientName = strPrefix & " " & Ucase_string(FirstName) & " " & Ucase_string(MiddleName) & " " & Ucase_string(LastName)
.Sex = StrSex
.Age = inAge
.Prefix_To_name = strPrefix
End With

End Function



















