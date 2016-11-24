VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Clinic Records"
   ClientHeight    =   10830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14730
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File_menu 
      Caption         =   "File"
      Begin VB.Menu Free_Registretation_Menu 
         Caption         =   "Free Registretation"
      End
      Begin VB.Menu LearnOnly_menu 
         Caption         =   "LearnOnly Heart"
      End
      Begin VB.Menu Forum_menu 
         Caption         =   "Forum"
      End
      Begin VB.Menu Printer_Menu 
         Caption         =   "Printer"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Appointment_menu 
      Caption         =   "&Appointment"
   End
   Begin VB.Menu Setting_menu 
      Caption         =   "Setting"
   End
   Begin VB.Menu Patient_details_menu 
      Caption         =   "Patient &Details"
   End
   Begin VB.Menu New_Doctor_menu 
      Caption         =   "New Doctor"
   End
   Begin VB.Menu New_drug_menu 
      Caption         =   "New Drug or Test"
   End
   Begin VB.Menu SMS_Menu 
      Caption         =   "S&MS"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function OpenURL(ByVal URL As String) As Long
    OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub Appointment_menu_Click()
FrmAppointment.Show
End Sub

Private Sub Forum_menu_Click()
OpenURL ("http://www.learnonly.com/p/learnonly-heart.html#.T-3JeRee4bA")
End Sub

Private Sub Free_Registretation_Menu_Click()
OpenURL ("http://feedburner.google.com/fb/a/mailverify?uri=learnonly/Apps")
End Sub

Private Sub LearnOnly_menu_Click()
OpenURL ("http://www.learnonly.com")
End Sub

Private Sub MDIForm_Load()
If LabMode = True Then
New_drug_menu.Caption = "New Test"
End If
End Sub




Private Sub New_Doctor_menu_Click()
FrmNewDoctor.Show
End Sub

Private Sub New_drug_menu_Click()
FrmNewMedicne.Show
End Sub

Private Sub Patient_details_menu_Click()
frmPatient.Show
End Sub

Private Sub Printer_Menu_Click()
CommonDialog.ShowPrinter
End Sub

Private Sub Setting_menu_Click()
FrmSetting.Show
End Sub

Private Sub SMS_Menu_Click()
frmSendSMS.Show
End Sub
