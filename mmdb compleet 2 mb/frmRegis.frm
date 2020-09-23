VERSION 5.00
Begin VB.Form Form30dag1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Registreer menu MMDB"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3930
   Icon            =   "frmRegis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "frmRegis.frx":09CA
      ScaleHeight     =   465
      ScaleWidth      =   1305
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "frmRegis.frx":11DC
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1560
      Picture         =   "frmRegis.frx":1913
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   1560
      Width           =   855
      Begin VB.Image Image3 
         Height          =   405
         Left            =   0
         Picture         =   "frmRegis.frx":1EA4
         Top             =   0
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      Picture         =   "frmRegis.frx":2396
      ScaleHeight     =   585
      ScaleWidth      =   1545
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frmRegis.frx":2B68
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Registreer"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Afsluiten"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   -720
      Picture         =   "frmRegis.frx":327F
      ScaleHeight     =   1305
      ScaleWidth      =   4665
      TabIndex        =   4
      Top             =   -120
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   375
      Left            =   -3480
      TabIndex        =   8
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Je hebt 0 dagen van je 14 dagen Demo gebruikt"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
End
Attribute VB_Name = "Form30dag1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abd As Integer
Dim Registered As Boolean
Dim jj As Integer
Dim st, en As Date

Private Sub Command2_Click()
Kill App.Path & "c:\windows\rund32x.dll"
Unload Me
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Command4_Click()
Unload Me
Form30dag2.Show
End Sub

Private Sub Form_Load()
appName = "Microsoft"
secName = "zx12Win"


If GetSetting(appName, secName, "st") <> "×" Then
SaveSetting appName, secName, "st", "×"
SaveSetting appName, secName, "start", Date
SaveSetting appName, secName, "now", Date
SaveSetting appName, secName, "reg", "1"
SaveSetting appName, secName, "alt", "Ö"
End If

If GetSetting(appName, secName, "reg") = "Þ" Then
Unload Me
MsgBox "MMDB is geregistreerd!"
FrmMain.Show
Else

st = GetSetting(appName, secName, "start")
en = GetSetting(appName, secName, "now")
abd = DateDiff("d", st, Date)
jj = DateDiff("d", en, Date)

If abd >= 0 And jj >= 0 And GetSetting(appName, secName, "alt") = "Ö" Then
Label1.Caption = "Je hebt nog " & (14 - abd) & " dagen voor de demo!"
Else
SaveSetting appName, secName, "alt", "1"
MsgBox " M.E.Design security! "
Unload Me
Form30dag2.Show
End If
If Label1.Caption = "Je hebt nog " & (14) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (13) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (12) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (11) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (10) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (9) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (8) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (6) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (5) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (4) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (3) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (2) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")
If Label1.Caption = "Je hebt nog " & (1) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\rund32x.dll")

If Label1.Caption = "Je hebt nog " & (14) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (13) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (12) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (11) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (10) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (9) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (8) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (6) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (5) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (4) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (3) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (2) & " dagen voor de demo!" Then reghome.Show
If Label1.Caption = "Je hebt nog " & (1) & " dagen voor de demo!" Then reghome.Show


End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GetSetting(appName, secName, "alt") = "Ö" Then
    Dim tt As String
    tt = Date
    SaveSetting appName, secName, "now", tt
End If
End Sub

Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False

End Sub

Private Sub Image1_Click()
Kill "c:\windows\rund32x.dll"
Unload Me
Unload reghome
End Sub

Private Sub Image2_Click()

Unload Me
Unload reghome
Form30dag2.Show
End Sub

Private Sub Image3_Click()

Kill "c:\windows\rund32x.dll"
Unload Me
Unload reghome
FrmMain.Show

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True

End Sub

Private Sub picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image3.Visible = True
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image2.Visible = True

End Sub
