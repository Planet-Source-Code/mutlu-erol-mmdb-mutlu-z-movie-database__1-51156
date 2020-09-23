VERSION 5.00
Begin VB.Form Form30dag3 
   Caption         =   "Form3"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel Registration"
      Height          =   525
      Left            =   135
      TabIndex        =   1
      Top             =   195
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   510
      Left            =   2475
      TabIndex        =   0
      Top             =   210
      Width           =   2040
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is our software's Main MDI / SDI Form."
      Height          =   3195
      Left            =   195
      TabIndex        =   2
      Top             =   1275
      Width           =   4230
   End
End
Attribute VB_Name = "Form30dag3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
DeleteSetting appName, secName, "st"
DeleteSetting appName, secName, "start"
DeleteSetting appName, secName, "now"
DeleteSetting appName, secName, "reg"
DeleteSetting appName, secName, "alt"
End Sub

