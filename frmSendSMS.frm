VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSendSMS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send SMS"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAll 
      Caption         =   "SMS to All"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox ComDoctor 
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
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3840
      Width           =   5415
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtNumber 
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
      Left            =   1800
      TabIndex        =   1
      Text            =   "+91"
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Default         =   -1  'True
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
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label LabTextLength 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label LabDoctor 
      Caption         =   "Doctor name"
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
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Message:"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Recipient (Phone number): with National Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmSendSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DocMobile As New Collection
Private strBuffer As String 'receive buffer

'Check for success
Private Function IsSuccess(ByVal data As String)
    IsSuccess = InStr(data, vbCrLf & "OK" & vbCrLf) > 0
End Function

'Check for comm error
Private Function IsCommError(ByVal data As String)
    IsCommError = InStr(data, vbCrLf & "ERROR" & vbCrLf) > 0
End Function

'Check for network error
Private Function IsNetworkError(ByVal data As String)
    IsNetworkError = InStr(data, vbCrLf & "+CMS ERROR:") > 0
End Function

'Check for any known error
Private Function IsError(ByVal data As String)
    IsError = IsCommError(data) Or IsNetworkError(data)
End Function

'Call this function when response is not the expected one.
'It analyzes the response and displays an appropriate error message if showMessage is True.
Private Sub ErrorDetails(ByVal data As String, Optional ByVal showMessage As Boolean = False)
    Dim msg As String
    
    'Check if there is data at all
    If Len(data) = 0 Then
        msg = "No answer from phone"
    Else
        'Check if comm error
        If IsCommError(data) Then
            msg = "Phone returned an error."
        Else
            'Check if network error
            If IsNetworkError(data) Then
                msg = "A network error occurred."
            Else
                'Anything else: Unexpected
                msg = "Unexpected response: """ & data & """"
            End If
        End If
    End If
    Call Trace(msg)
    If showMessage Then Call MsgBox(msg, vbCritical + vbOKOnly)
End Sub

'Ensures that the string contains a success message,
'If not, determines the error details.
Private Function VerifySuccess(ByVal data As String, Optional ByVal showMessage As Boolean = False) As Boolean
    VerifySuccess = True
    If Not IsSuccess(data) Then
        VerifySuccess = False
        Call ErrorDetails(data, showMessage)
    End If
End Function

'Ensures that the string ends with a specific string.
'If not, determines the error details.
Private Function VerifyEndsWith(ByVal data As String, ByVal endsWith As String, Optional ByVal showMessage As Boolean = False) As Boolean
    VerifyEndsWith = True
    If InStr(data, endsWith) <> (Len(data) - Len(endsWith) + 1) Then
        VerifyEndsWith = False
        Call ErrorDetails(data, showMessage)
    End If
End Function

'Sends data to the serial port
Private Sub Send(ByVal data As String)
    strBuffer = ""
    Call Trace("<< " & data)
    MSComm1.Output = data
End Sub

'Receives data from the serial port
Private Function Receive() As String
    Dim strPart As String
    Dim strInput As String
    strInput = ""
    Do
        strPart = ""
        Call Delay(1)
        strPart = strBuffer
        strBuffer = ""
        If strPart = "" Then Exit Do
        strInput = strInput & strPart
    Loop
    If strInput <> "" Then Call Trace(">> " & strInput)
    Receive = strInput
End Function

'Waits until a success message is received or a timeout occurred.
Private Function WaitForSuccess(ByRef data As String) As Boolean
    Dim I As Integer
    Dim strInput As String
    Dim strPart As String
    strInput = ""
   
    'try receive 5 times with 2 seconds delay between
    For I = 1 To 5
        strPart = Receive
        strInput = strInput & strPart
        If IsSuccess(strInput) Or IsError(strInput) Then Exit For
        If strPart = "" Then
            Call Trace("Nothing new received, waiting 2s...")
            Call Delay(2)
        End If
    Next
    data = strInput
    WaitForSuccess = IsSuccess(strInput)
End Function

Private Sub Trace(ByVal message As String)
    Dim strLine As String
    strLine = DateTime.Now & " " & message
    txtLog.Text = txtLog.Text & strLine & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub btnSend_Click()
    txtLog.Text = ""
    
    txtNumber.Enabled = False
    txtMessage.Enabled = False
    btnSend.Enabled = False
    btnQuit.Enabled = False
    
    On Error GoTo ErrorHandler
    With MSComm1
        .CommPort = intComport
        .Settings = "19200,N,8,1"
        .Handshaking = comRTS
        .RTSEnable = True
        .DTREnable = True
        .RThreshold = 1
        .SThreshold = 1
        .InputMode = comInputModeBinary
        .InputLen = 0
        .PortOpen = True 'must be the last
    End With
    
    
    'Call Send(Chr(27)) 'Cancel active message input
    
    'Verify connection
    Call Send("AT" & vbCrLf)
    If Not VerifySuccess(Receive, True) Then GoTo finally
    
    'Enter text mode (phone must support it)
    Call Send("AT+CMGF=1" & vbCrLf)
    If Not VerifySuccess(Receive, True) Then GoTo finally
    
    'Send the message
    Dim buf As String
    Call Send("AT+CMGS=" & Chr(34) & txtNumber.Text & Chr(34) & vbCrLf)
    If Not VerifyEndsWith(Receive, "> ", True) Then GoTo finally
    Call Send(txtMessage.Text & " " & SMSuffix & Chr(26))
    If WaitForSuccess(buf) Then
        Call MsgBox("Message was submitted successfully.", vbInformation + vbOKOnly)
    Else
        Call ErrorDetails(buf, True)
    End If
    
    GoTo finally
ErrorHandler:
    Call MsgBox("Error: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)
finally:
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    txtNumber.Enabled = True
    txtMessage.Enabled = True
    btnSend.Enabled = True
    btnQuit.Enabled = True
End Sub






Private Sub CmdAll_Click()
txtLog.Text = ""
    
    txtNumber.Enabled = False
    txtMessage.Enabled = False
    btnSend.Enabled = False
    btnQuit.Enabled = False
    
    On Error GoTo ErrorHandler
    With MSComm1
        .CommPort = intComport
        .Settings = "19200,N,8,1"
        .Handshaking = comRTS
        .RTSEnable = True
        .DTREnable = True
        .RThreshold = 1
        .SThreshold = 1
        .InputMode = comInputModeBinary
        .InputLen = 0
        .PortOpen = True 'must be the last
    End With
    
    
    'Call Send(Chr(27)) 'Cancel active message input
    
    'Verify connection
    Call Send("AT" & vbCrLf)
    If Not VerifySuccess(Receive, True) Then GoTo finally
    
    'Enter text mode (phone must support it)
    Call Send("AT+CMGF=1" & vbCrLf)
    If Not VerifySuccess(Receive, True) Then GoTo finally
    
    'Send the message
    Dim buf As String
    Call Send("AT+CMGS=" & Chr(34) & txtNumber.Text & Chr(34) & vbCrLf)
    If Not VerifyEndsWith(Receive, "> ", True) Then GoTo finally
    Call Send(txtMessage.Text & " " & SMSuffix & Chr(26))
    If WaitForSuccess(buf) Then
        Call MsgBox("Message was submitted successfully.", vbInformation + vbOKOnly)
    Else
        Call ErrorDetails(buf, True)
    End If
    
    GoTo finally
ErrorHandler:
    Call MsgBox("Error: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)
finally:
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    txtNumber.Enabled = True
    txtMessage.Enabled = True
    btnSend.Enabled = True
    btnQuit.Enabled = True
End Sub

Private Sub ComDoctor_LostFocus()
On Error GoTo trap
txtNumber.Text = CountryCode & DocMobile(ComDoctor.ListIndex + 1)
Exit Sub
trap:

End Sub

Private Sub Form_Load()
 Dim strSql As String
Dim CurPopulate As New ADODB.Recordset

    txtNumber.SelStart = 0
    txtNumber.SelLength = Len(txtNumber.Text)
    
   


strSql = "SELECT Doctor_details.Doctor_Name, Doctor_details.Mobile " & _
"From Doctor_details " & _
"ORDER BY Doctor_details.Doctor_Name;"


With CurPopulate
.Open strSql, CnnMain, adOpenStatic, adLockReadOnly, adCmdText
If .RecordCount = 0 Then
Exit Sub
End If
If .EOF = False Then
.MoveLast
.MoveFirst
Do While Not .EOF
        ComDoctor.AddItem (.Fields("Doctor_Name"))
        DocMobile.Add (.Fields("Mobile").Value)
               
        .MoveNext
    Loop
    End If
    .Close
    
End With

    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not btnQuit.Enabled Then Cancel = True
End Sub

Private Sub btnQuit_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)


Set DocMobile = Nothing



End Sub

Private Sub MSComm1_OnComm()
    Dim strMessage As String
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            strMessage = StrConv(MSComm1.Input, vbUnicode)
           
            
'        Case comEvSend
'        Case comEvCTS
'            strMessage = "[Change in CTS Detected]"
'        Case comEvDSR
'            strMessage = "[Change in DSR Detected]"
'        Case comEvCD
'            strMessage = "[Change in CD Detected]"
'        Case comEvRing
'            strMessage = "[The Phone is Ringing]"
'        Case comEvEOF
'            strMessage = "[End of File Detected]"
            
        ' Error messages.
        Case comBreak
            strMessage = "[Break Received]"
        Case comCDTO
            strMessage = "[Carrier Detect Timeout]"
        Case comCTSTO
            strMessage = "[CTS Timeout]"
        Case comDCB
            strMessage = "[Error retrieving DCB]"
        Case comDSRTO
            strMessage = "[DSR Timeout]"
        Case comFrame
            strMessage = "[Framing Error]"
        Case comOverrun
            strMessage = "[Overrun Error]"
        Case comRxOver
            strMessage = "[Receive Buffer Overflow]"
        Case comRxParity
            strMessage = "[Parity Error]"
        Case comTxFull
            strMessage = "[Transmit Buffer Full]"
'        Case Else
'            strMessage = "[Unknown error or event: " & MSComm1.CommEvent & "]"
    End Select
    strBuffer = strBuffer & strMessage
End Sub

Private Sub Delay(ByVal HowLong As Date)
    Dim endDate As Date
    endDate = DateAdd("s", HowLong, Now)
    While endDate > Now
        DoEvents 'Allows windows to handle other stuff
    Wend
End Sub

Private Sub txtMessage_Change()
LabTextLength.Caption = Len(txtMessage.Text)
End Sub
