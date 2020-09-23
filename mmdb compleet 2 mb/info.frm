VERSION 5.00
Begin VB.Form info 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Start, Programma's,Mutlu'z movie database, start  =>MMDB acteurdatabase "
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
