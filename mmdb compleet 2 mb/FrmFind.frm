VERSION 5.00
Begin VB.Form FrmFind 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Film zoeken"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "FrmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5550
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3840
      Picture         =   "FrmFind.frx":09CA
      ScaleHeight     =   705
      ScaleWidth      =   1545
      TabIndex        =   3
      Top             =   960
      Width           =   1575
      Begin VB.Image Image2 
         Height          =   765
         Left            =   0
         Picture         =   "FrmFind.frx":12C4
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3840
      Picture         =   "FrmFind.frx":1DA8
      ScaleHeight     =   705
      ScaleWidth      =   1545
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      Begin VB.Image Image1 
         Height          =   765
         Left            =   0
         Picture         =   "FrmFind.frx":26DD
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.TextBox findstr 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "Zoeken"
      Height          =   195
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Typ de naam van de film om te zoeken:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldfindstart As Integer
Private Sub Command1_Click()
If oldfindstart = findstart And findstart > 0 Then
findstart = 0
oldfindstart = 0
Call find(Trim(findstr.Text), findstart)
Exit Sub
End If
oldfindstart = findstart
Call find(Trim(findstr.Text), findstart)

End Sub


Private Sub findstr_Change()
findstart = 0

End Sub

Private Sub findstr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub



Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Image2.Visible = False

End Sub

Private Sub Image1_Click()

If oldfindstart = findstart And findstart > 0 Then
findstart = 0
oldfindstart = 0
Call find(Trim(findstr.Text), findstart)
Exit Sub
End If
oldfindstart = findstart
Call find(Trim(findstr.Text), findstart)

End Sub

Private Sub Image2_Click()
Unload Me

End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image2.Visible = True
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True

End Sub
