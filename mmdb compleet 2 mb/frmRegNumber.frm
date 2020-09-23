VERSION 5.00
Begin VB.Form Form30dag2 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Registreer MMDB"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3195
   Icon            =   "frmRegNumber.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      Picture         =   "frmRegNumber.frx":09CA
      ScaleHeight     =   465
      ScaleWidth      =   1425
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "frmRegNumber.frx":119C
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      Picture         =   "frmRegNumber.frx":18B3
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   960
      Width           =   855
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frmRegNumber.frx":1E44
         Top             =   0
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Afsluiten"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   99
      PasswordChar    =   "@"
      TabIndex        =   0
      Top             =   600
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "mmdb01012003  ||||serienummer"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Vul de serienummer van MMDB"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form30dag2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "mmdb01012003" Then
SaveSetting appName, secName, "reg", "Þ"
Unload Me
FrmMain.Show
End If
End Sub

Private Sub Command2_Click()
Kill "c:\windows\rund32x.dll"
Unload Me
End Sub

Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Image2.Visible = False
End Sub

Private Sub Image1_Click()
If Text1 = "mmdb01012003" Then
SaveSetting appName, secName, "reg", "Þ"
Unload Me
FrmMain.Show
End If
End Sub

Private Sub Image2_Click()
Kill "c:\windows\rund32x.dll"
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image2.Visible = True

End Sub
