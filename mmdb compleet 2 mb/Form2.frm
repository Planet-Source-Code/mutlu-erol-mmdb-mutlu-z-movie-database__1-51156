VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   480
      Picture         =   "Form2.frx":09CA
      ScaleHeight     =   1065
      ScaleWidth      =   4785
      TabIndex        =   9
      Top             =   360
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2400
      Picture         =   "Form2.frx":2B0F
      ScaleHeight     =   705
      ScaleWidth      =   1545
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
      Begin VB.Image Image1 
         Height          =   765
         Left            =   0
         Picture         =   "Form2.frx":3409
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2000
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   1080
         TabIndex        =   2
         Top             =   1560
         Width           =   4215
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form2.frx":3EED
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   2955
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   3480
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form2.frx":40C7
            ForeColor       =   &H0080C0FF&
            Height          =   960
            Left            =   240
            TabIndex        =   6
            Top             =   3360
            Width           =   3360
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Als je een idee hebt wat ik in mijn programma kan toevoegen, mail me effe, zodat ik deze programma kan verbeteren........."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   915
            Left            =   480
            TabIndex        =   5
            Top             =   7440
            Width           =   3240
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            X1              =   240
            X2              =   3720
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404040&
            X1              =   240
            X2              =   3720
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Voor vragen, suggesties of informatie kun je me mailen op deze mailadres:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   675
            Left            =   720
            TabIndex        =   4
            Top             =   4800
            Width           =   2880
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "tatlicocuk21@hotmail.com"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   675
            Left            =   960
            TabIndex        =   3
            Top             =   5640
            Width           =   2295
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00404040&
            X1              =   360
            X2              =   3840
            Y1              =   6720
            Y2              =   6720
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00404040&
            X1              =   960
            X2              =   3120
            Y1              =   6840
            Y2              =   6840
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00404040&
            X1              =   1200
            X2              =   2880
            Y1              =   6960
            Y2              =   6960
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00404040&
            X1              =   1440
            X2              =   2640
            Y1              =   7080
            Y2              =   7080
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00404040&
            X1              =   1680
            X2              =   2400
            Y1              =   7200
            Y2              =   7200
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00404040&
            X1              =   1920
            X2              =   2160
            Y1              =   7320
            Y2              =   7320
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5535
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2002 Mutlu Erol, Inc. All rights reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   4680
      Width           =   4140
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False

End Sub

Private Sub Image1_Click()

Unload Me

End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True

End Sub

Private Sub Timer1_Timer()
Frame2.Top = Frame2.Top - 20
If Frame2.Top <= -7200 Then
    Frame2.Top = 1680
End If
End Sub
