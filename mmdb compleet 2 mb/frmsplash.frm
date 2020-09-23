VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MousePointer    =   11  'Hourglass
   Picture         =   "frmsplash.frx":09CA
   ScaleHeight     =   5355
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   2000
      Picture         =   "frmsplash.frx":39E2
      ScaleHeight     =   1470
      ScaleWidth      =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   3150
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tatlicocuk21@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2520
      TabIndex        =   6
      Top             =   4560
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "       voor info of vragen kun je mailen naar:      "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   1680
      TabIndex        =   5
      Top             =   4200
      Width           =   3660
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   1000
      Width           =   5295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bezig met laden ... . . . . . . . "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2640
      TabIndex        =   3
      Top             =   2745
      Width           =   2130
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   6240
      X2              =   6240
      Y1              =   3000
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   720
      X2              =   720
      Y1              =   3000
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6960
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmsplash.frx":5B27
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2002.Mutlu Erol, Inc. All right reserved."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   3750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mutlu'z Movie DBaSe"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   3300
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
Load FrmMain
FrmMain.Show
End Sub
