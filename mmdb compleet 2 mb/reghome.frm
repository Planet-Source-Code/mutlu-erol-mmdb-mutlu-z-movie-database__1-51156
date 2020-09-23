VERSION 5.00
Begin VB.Form reghome 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "7 daags test periode MMDB"
   ClientHeight    =   3345
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   6540
   Icon            =   "reghome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   720
      Picture         =   "reghome.frx":09CA
      ScaleHeight     =   825
      ScaleWidth      =   4785
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   6360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   6360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"reghome.frx":2B0F
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
End
Attribute VB_Name = "reghome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\chill.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\database.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\finish.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\mmdb.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\mutlu.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\relax.dll")
'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Call FileCopy(App.Path & "\reginfo.txt", "c:\windows\startxl.dll")

'If Label1.Caption = "Je hebt nog " & (7) & " dagen voor de demo!" Then Kill App.Path & "\chill.dll"
'If Label1.Caption = "Je hebt nog " & (6) & " dagen voor de demo!" Then Kill App.Path & "\database.dll"
'If Label1.Caption = "Je hebt nog " & (5) & " dagen voor de demo!" Then Kill App.Path & "\finish.dll"
'If Label1.Caption = "Je hebt nog " & (4) & " dagen voor de demo!" Then Kill App.Path & "\mmdb.dll"
'If Label1.Caption = "Je hebt nog " & (3) & " dagen voor de demo!" Then Kill App.Path & "\mutlu.dll"
'If Label1.Caption = "Je hebt nog " & (2) & " dagen voor de demo!" Then Kill App.Path & "\relax.dll"
'If Label1.Caption = "Je hebt nog " & (1) & " dagen voor de demo!" Then Kill App.Path & "\startxl.dll"


End Sub

Private Sub Command2_Click()
Dim frmB As New frmBrowser
frmB.StartingAddress = "http://members.home.nl/m.erol/update.htm"
frmB.Show
End Sub
