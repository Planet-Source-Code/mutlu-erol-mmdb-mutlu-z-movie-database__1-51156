VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   LinkTopic       =   "Form5"
   Moveable        =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Je hebt MMDB al 7 dagen gebruikt"
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

