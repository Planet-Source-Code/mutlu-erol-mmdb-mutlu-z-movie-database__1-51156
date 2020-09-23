VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutlu'z Movie DataBase"
   ClientHeight    =   10470
   ClientLeft      =   90
   ClientTop       =   765
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   472.686
   ScaleMode       =   0  'User
   ScaleWidth      =   465.792
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   9000
      Width           =   15255
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2280
         Picture         =   "FrmMain.frx":09CA
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   38
         Top             =   120
         Width           =   1575
         Begin VB.Image Image10 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":13B2
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2280
         Picture         =   "FrmMain.frx":1F67
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   97
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Image Image11 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":28F4
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         Picture         =   "FrmMain.frx":340C
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   39
         Top             =   120
         Width           =   1575
         Begin VB.Image Image8 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":3E22
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         Picture         =   "FrmMain.frx":4A63
         ScaleHeight     =   825
         ScaleWidth      =   1545
         TabIndex        =   95
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Image Image9 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":53F0
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   735
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Afsluiten"
         Height          =   195
         Left            =   14280
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6240
         Picture         =   "FrmMain.frx":5F08
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   46
         Top             =   120
         Width           =   1575
         Begin VB.Image Image6 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":672A
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4320
         Picture         =   "FrmMain.frx":71B4
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   12
         Top             =   120
         Width           =   1575
         Begin VB.Image Image7 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":7C40
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   8040
         Picture         =   "FrmMain.frx":882E
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   37
         Top             =   120
         Width           =   1575
         Begin VB.Image Image5 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":9207
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9840
         Picture         =   "FrmMain.frx":9D85
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   36
         Top             =   120
         Width           =   1575
         Begin VB.Image Image4 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":A82C
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000A&
         Caption         =   "&Verwijder"
         Height          =   255
         Left            =   14640
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Zoeken"
         Height          =   195
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command11 
         Caption         =   "filmlijst"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ok"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   135
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Verwijder alles"
         Height          =   195
         Left            =   13680
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   11640
         Picture         =   "FrmMain.frx":B439
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   40
         Top             =   120
         Width           =   1575
         Begin VB.Image Image3 
            Height          =   765
            Left            =   0
            Picture         =   "FrmMain.frx":BD6E
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.Image Image2 
         Height          =   765
         Left            =   13440
         Picture         =   "FrmMain.frx":C883
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   765
         Left            =   13440
         Picture         =   "FrmMain.frx":D426
         Top             =   120
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   15480
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8760
         TabIndex        =   115
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Zoek"
         Height          =   255
         Left            =   8160
         TabIndex        =   114
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "FrmMain.frx":DE19
         ScaleHeight     =   585
         ScaleWidth      =   1065
         TabIndex        =   113
         Top             =   240
         Width           =   1095
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "FrmMain.frx":EA24
         Height          =   6300
         Left            =   12240
         TabIndex        =   112
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   11113
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8421504
         ForeColor       =   4194304
         ListField       =   "filmnaam"
         Object.DataMember      =   "connect"
      End
      Begin VB.TextBox txtfilmnaam 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5880
         Width           =   2775
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4455
         Left            =   480
         ScaleHeight     =   4395
         ScaleWidth      =   2955
         TabIndex        =   34
         Top             =   1680
         Width           =   3015
         Begin VB.PictureBox picpic 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2040
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   35
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            Caption         =   "resolutie 193 * 258 "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   117
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H00000000&
            Caption         =   "Klik op Cover selecteren"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   116
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            Caption         =   "                   [MMDB 2003 ©]"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6120
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton Option1 
            BackColor       =   &H00400000&
            Caption         =   "DivX"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00400000&
            Caption         =   "DVD"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   72
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Textkeuze 
            Height          =   375
            Left            =   3000
            TabIndex        =   71
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00400000&
            Caption         =   "VHS"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   70
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00400000&
            Caption         =   "VCD"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   69
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00400000&
            Caption         =   "SVCD"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   68
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00400000&
            Caption         =   "PDivX"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   67
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtgenre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6120
         TabIndex        =   6
         Text            =   "Actie ,horror,thriller.etx"
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox txtcd 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6120
         TabIndex        =   1
         Text            =   "Vul hier de naam van de film"
         Top             =   1680
         Width           =   4335
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         Picture         =   "FrmMain.frx":EA43
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   42
         Top             =   6360
         Width           =   375
         Begin VB.Image Image14 
            Height          =   375
            Left            =   0
            Picture         =   "FrmMain.frx":EDAE
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         MouseIcon       =   "FrmMain.frx":F11A
         Picture         =   "FrmMain.frx":F424
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   41
         Top             =   6360
         Width           =   375
         Begin VB.Image Image15 
            Height          =   375
            Left            =   0
            Picture         =   "FrmMain.frx":F791
            Top             =   0
            Width           =   375
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9360
         Top             =   6840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   8520
         Top             =   7200
      End
      Begin VB.TextBox txtacteurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Vul hier de naam van de acteurs"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txttype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "DivX of DVD of VHS"
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox txtspeelduur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Vul hier de speelduur van de film"
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtrating 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "**********"
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox txtuitgeleend 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "Aan wie heb je hem uitgeleend?"
         Top             =   5280
         Width           =   2895
      End
      Begin VB.TextBox txtmisc 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "FrmMain.frx":FB17
         Top             =   7680
         Width           =   6255
      End
      Begin VB.TextBox txtjaar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Welke jaar?"
         Top             =   3840
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000012&
         Caption         =   "<<"
         Height          =   105
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         Top             =   5040
         Width           =   135
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000012&
         Caption         =   ">>"
         Height          =   105
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   18
         Top             =   4320
         Width           =   135
      End
      Begin VB.TextBox txtpost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1455
         TabIndex        =   17
         Text            =   "1"
         Top             =   6360
         Width           =   960
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   11055
         Left            =   120
         Picture         =   "FrmMain.frx":FB31
         ScaleHeight     =   11025
         ScaleWidth      =   14985
         TabIndex        =   50
         Top             =   -120
         Width           =   15015
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   12720
            TabIndex        =   120
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Zoeken"
            Height          =   255
            Left            =   13680
            TabIndex        =   121
            Top             =   720
            Width           =   975
         End
         Begin VB.PictureBox Picture18 
            Height          =   375
            Left            =   2880
            Picture         =   "FrmMain.frx":3DD8C
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   119
            Top             =   6480
            Width           =   375
         End
         Begin VB.PictureBox Picture17 
            Height          =   375
            Left            =   360
            Picture         =   "FrmMain.frx":3E10A
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   118
            Top             =   6480
            Width           =   375
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Help"
            ForeColor       =   &H80000008&
            Height          =   3375
            Left            =   10080
            TabIndex        =   103
            Top             =   2880
            Visible         =   0   'False
            Width           =   2655
            Begin VB.TextBox Text7 
               Height          =   615
               Left            =   240
               MultiLine       =   -1  'True
               TabIndex        =   111
               Text            =   "FrmMain.frx":3E46C
               Top             =   2640
               Width           =   1695
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   240
               TabIndex        =   110
               Text            =   "**********"
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   240
               TabIndex        =   109
               Text            =   "-"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   240
               TabIndex        =   108
               Text            =   "Actie"
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   240
               TabIndex        =   107
               Text            =   "    minuten"
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   240
               TabIndex        =   106
               Text            =   "2002"
               Top             =   1200
               Width           =   855
            End
            Begin VB.CommandButton Command40 
               Caption         =   "on"
               Height          =   255
               Left            =   2160
               TabIndex        =   105
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   240
               TabIndex        =   104
               Text            =   "DivX"
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Film info op IMDB.com"
            Height          =   375
            Left            =   720
            TabIndex        =   102
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton Command38 
            BackColor       =   &H00000000&
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   101
            Top             =   5400
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command37 
            Caption         =   ">>"
            Height          =   255
            Left            =   10560
            TabIndex        =   100
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command36 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   99
            Top             =   2520
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command29 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   93
            Top             =   4920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command30 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   94
            Top             =   4920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6480
            TabIndex        =   90
            Top             =   4800
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton Option8 
               BackColor       =   &H00400000&
               Caption         =   "Nee"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   960
               TabIndex        =   92
               Top             =   120
               Width           =   975
            End
            Begin VB.OptionButton Option7 
               BackColor       =   &H00400000&
               Caption         =   "Ja"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.TextBox ondertiteling 
            Appearance      =   0  'Flat
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   6000
            TabIndex        =   7
            Text            =   "Ondertiteling?"
            Top             =   4920
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CommandButton Command14 
            Caption         =   ">>"
            Height          =   255
            Left            =   2640
            TabIndex        =   88
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command23 
            Caption         =   "<<"
            Height          =   255
            Left            =   2400
            TabIndex        =   89
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cds 
            BackColor       =   &H00400000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   87
            Text            =   "kies!"
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command13 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   54
            Top             =   5880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command15 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   55
            Top             =   4440
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command16 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   56
            Top             =   3960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command17 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   57
            Top             =   3480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command20 
            Caption         =   ">>"
            Height          =   255
            Left            =   9000
            TabIndex        =   58
            Top             =   3000
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   6000
            TabIndex        =   75
            Top             =   5880
            Visible         =   0   'False
            Width           =   2895
            Begin VB.CheckBox Check1 
               BackColor       =   &H00400000&
               Caption         =   "Ik vond het niks aan"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   86
               Top             =   0
               Width           =   2175
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00400000&
               Caption         =   "Valt niet weer te kijken"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   85
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00400000&
               Caption         =   "Ging wel"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   84
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox Check4 
               BackColor       =   &H00400000&
               Caption         =   "Idee is goed, maar slecht geacteerd"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   83
               Top             =   720
               Width           =   3015
            End
            Begin VB.CheckBox Check5 
               BackColor       =   &H00400000&
               Caption         =   "Eeen dood gewone film"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   82
               Top             =   960
               Width           =   2295
            End
            Begin VB.CheckBox Check6 
               BackColor       =   &H00400000&
               Caption         =   "Voldoende"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   81
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CheckBox Check7 
               BackColor       =   &H00400000&
               Caption         =   "Leuke film"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   80
               Top             =   1440
               Width           =   2295
            End
            Begin VB.CheckBox Check8 
               BackColor       =   &H00400000&
               Caption         =   "Goeie film"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   79
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox Check9 
               BackColor       =   &H00400000&
               Caption         =   "Strakke film"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   78
               Top             =   1920
               Width           =   2415
            End
            Begin VB.CheckBox Check10 
               BackColor       =   &H00400000&
               Caption         =   "Super film "
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   77
               Top             =   2160
               Width           =   3495
            End
            Begin VB.TextBox ratingbalk 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   76
               Top             =   2520
               Width           =   2055
            End
         End
         Begin VB.ComboBox minuten 
            BackColor       =   &H00400000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6000
            TabIndex        =   74
            Text            =   "Aantal minuten"
            Top             =   3000
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.ComboBox Jaar 
            BackColor       =   &H00400000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6000
            TabIndex        =   65
            Text            =   "Welke jaar?"
            Top             =   3960
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CommandButton Command28 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   64
            Top             =   5880
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command27 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   63
            Top             =   3000
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command26 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   62
            Top             =   3480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command25 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   61
            Top             =   3960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command24 
            Caption         =   "<<"
            Height          =   255
            Left            =   9000
            TabIndex        =   60
            Top             =   4440
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox keuzemenu 
            BackColor       =   &H00400000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6000
            TabIndex        =   59
            Text            =   "Kies genre"
            Top             =   4440
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox aantalcds 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1320
            TabIndex        =   11
            Text            =   "cd's?"
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Filmlijst HTML converter"
            Height          =   495
            Left            =   1080
            TabIndex        =   98
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Image Image13 
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image Image12 
            Height          =   855
            Left            =   -120
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   195
         Left            =   14280
         TabIndex        =   48
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   195
         Left            =   14280
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "&Nieuwe film"
         Height          =   315
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   8280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000A&
         Caption         =   "&Bewerken"
         Height          =   255
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   8640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Acteurs                    :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Genre                      :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Type film                  :            "
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   26
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "Speelduur                :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Rating  (max/10)     :"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         Caption         =   "Jaar                         :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Uitgeleend aan        :"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000007&
         Caption         =   "Films lijst"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   13200
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Film naam"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Film info"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   6840
         Width           =   1695
      End
      Begin VB.Label lblcreated 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   4200
         Width           =   3375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   10200
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18918
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   30
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bezig met comprimeren van de database.  Heb geduld............."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   495
         Left            =   480
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnAuto As Boolean
Dim editing As Boolean
Private Sub about_Click()
If frmAbout.Visible = False Then
frmAbout.Show 0, Me
Else
Unload frmAbout
End If
End Sub
Private Sub Afsluiten_Click()
If editing = True Then
    tmp = MsgBox("Weet je het zeker dat je wilt gaan stoppen?" & vbCrLf & "De wijzigingen worden niet gesaved!.", vbYesNo, "M.E.")
    If tmp = 7 Then Exit Sub
End If
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Picture1.Visible = True
Picture1.Refresh
End If
Unload FrmFind
Unload Me
End
End Sub
Private Sub Alles_Verwijderen_Click()
tmpo = MsgBox("Weet je het zeker om alle films te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Call CreateNewDB("mutlu.mdb")
Form_Load
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
ondertiteling = ""
End Sub
Private Sub Bewerken_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    End If
    editing = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    Command5.Caption = "&Opslaan"
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
Else
    editing = False
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Picture2.Cls
    updatedb (postnr)
    Command1.Enabled = True
    Command2.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    Command5.Caption = "&Bewerken"
End If
End Sub
Private Sub captureit_Click()
If Mutluz.Visible = False Then
Mutluz.Show 0, Me
Else
Unload Mutluz
End If
End Sub
Private Sub cboAuto_Change()
End Sub
Private Sub Command1_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    editing = True
    Command2.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtcd.SetFocus
    Command1.Caption = "&Opslaan"
    Else
    Picture2.Cls
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Command2.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    savetodb
    editing = False
    Command1.Caption = "&Nieuwe film"
End If
End Sub
Private Sub Command10_Click()
Kill App.Path & "\compact.mdb"
End Sub
Private Sub Command12_Click()
txtgenre = txtgenre
End Sub
Private Sub Command13_Click()
If txtrating.Visible = True Then Command28.Visible = True
If txtrating.Visible = True Then Command13.Visible = False
Frame4.Visible = True
txtrating.Visible = False
End Sub
Private Sub Command14_Click()
If aantalcds.Visible = True Then Command23.Visible = True
If aantalcds.Visible = True Then Command14.Visible = False
aantalcds.Visible = False
cds.Visible = True
aantalcds = cds
End Sub
Private Sub Command15_Click()
If txtgenre.Visible = True Then Command24.Visible = True
If txtgenre.Visible = True Then Command15.Visible = False
keuzemenu.Visible = True
txtgenre.Visible = False
 End Sub
Private Sub Command16_Click()
If txtjaar.Visible = True Then Command25.Visible = True
If txtjaar.Visible = True Then Command16.Visible = False
Jaar.Visible = True
txtjaar.Visible = False
End Sub
Private Sub Command17_Click()
If txttype.Visible = True Then Command26.Visible = True
If txttype.Visible = True Then Command17.Visible = False
Frame3.Visible = True
txttype.Visible = False
End Sub
Private Sub Command18_Click()
End Sub
Private Sub Command19_Click()
End Sub
Private Sub Command2_Click()
tmpo = MsgBox("Weet je het zeker om de geselecteerde film te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Dim tmp As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.Move postnr - 1
Rs.Delete
postnr = Rs.RecordCount
txtpost = postnr
If Rs.RecordCount > 0 Then
tmp = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost tmp
Open App.Path & "\mutlu.mdb" For Binary As #1
g = LOF(1)
Close #1
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Else
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtrating.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
    txtcd.Locked = True
    aantalcds.Locked = True
    ondertiteling.Locked = True
StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
Command8.Enabled = False
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End If
End Sub
Private Sub Command20_Click()
If txtspeelduur.Visible = True Then Command27.Visible = True
If txtspeelduur.Visible = True Then Command20.Visible = False
minuten.Visible = True
txtspeelduur.Visible = False
End Sub
Private Sub Command21_Click()
End Sub
Private Sub Command22_Click()
End Sub
Private Sub Command23_Click()
If cds.Visible = True Then Command23.Visible = False
If aantalcds.Visible = False Then Command14.Visible = True
cds.Visible = False
aantalcds.Visible = True
aantalcds = cds
End Sub
Private Sub Command24_Click()
If keuzemenu.Visible = True Then Command24.Visible = False
If txtgenre.Visible = False Then Command15.Visible = True
txtgenre.Visible = True
keuzemenu.Visible = False
txtgenre = keuzemenu
End Sub
Private Sub Command25_Click()
If Jaar.Visible = True Then Command25.Visible = False
If txtjaar.Visible = False Then Command16.Visible = True
txtjaar.Visible = True
Jaar.Visible = False
txtjaar = Jaar
End Sub
Private Sub Command26_Click()
If Frame3.Visible = True Then Command26.Visible = False
If txttype.Visible = False Then Command17.Visible = True
txttype.Visible = True
Frame3.Visible = False
End Sub
Private Sub Command27_Click()
If minuten.Visible = True Then Command27.Visible = False
If txtspeelduur.Visible = False Then Command20.Visible = True
txtspeelduur.Visible = True
minuten.Visible = False
txtspeelduur = minuten
End Sub
Private Sub Command28_Click()
If Frame4.Visible = True Then Command28.Visible = False
If txtrating.Visible = False Then Command13.Visible = True
txtrating.Visible = True
Frame4.Visible = False
End Sub
Private Sub Command29_Click()
If ondertiteling.Visible = True Then Command30.Visible = True
If ondertiteling.Visible = True Then Command29.Visible = False
Frame5.Visible = True
ondertiteling.Visible = False
End Sub
Private Sub Command3_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
End Sub
Private Sub Command31_Click()
Dim found As Boolean, t As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.MoveFirst
For t = start To Rs.RecordCount - 1
    With Rs
        For f = 0 To .Fields.Count - 1
            If .Fields(f) <> "" Then
            test = InStr(1, Text9, .Fields(f), vbTextCompare)
            If InStr(1, Text9, Trim(.Fields(f))) > 0 Then GoTo found
            End If
        Next f
        Rs.MoveNext
    End With
Next t
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Exit Sub
found:
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
findstart = t
visapost t + 1
postnr = t + 1
End Sub
Private Sub Command30_Click()
If Frame5.Visible = True Then Command30.Visible = False
If ondertiteling.Visible = False Then Command29.Visible = True
ondertiteling.Visible = True
Frame5.Visible = False
End Sub
Private Sub Command32_Click()
End Sub
Private Sub Command33_Click()
End Sub
Private Sub Command34_Click()
End Sub
Private Sub Command35_Click()
shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\check.bat", vbNullString, vbNullString, SW_Shownormal)
MsgBox "Vorige html-bestanden zijn gewist"
shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\kill.bat", vbNullString, vbNullString, SW_Shownormal)
MsgBox "Nieuwe Filmlijst html-bestand is aangemaakt "
shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\renamehtml.bat", vbNullString, vbNullString, SW_Shownormal)
MsgBox "Html converter word gestart"
Convert.Show
End Sub
Private Sub Command36_Click()
Dim f As New Acteurformulier
            f.Show
End Sub
Private Sub Command37_Click()
Dim f As New filmnaamformulier
            f.Show
End Sub
Private Sub Command38_Click()
Dim f As New klantenformulier
            f.Show
End Sub
Private Sub Command39_Click()
    Dim frmB As New frmBrowser
    frmB.StartingAddress = browserstring + "http://us.imdb.com/Tsearch?title=" + txtcd.Text + "&restrict=Movies"
    frmB.Show
End Sub
Private Sub Command4_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 1
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost postnr
End If
End Sub
Private Sub Command40_Click()
txttype = Text1
txtspeelduur = Text3
txtjaar = Text2
txtgenre = Text4
txtuitgeleend = Text5
txtrating = Text6
txtmisc = Text7
End Sub
Private Sub Command41_Click()
End Sub
Private Sub Command42_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = browserstring + "http://www.fasttrackmovies.com/listings.asp?L=" + Text8.Text + "&submit=Search"
    frmB.Show
End Sub
Private Sub Command5_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    End If
    editing = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    Command5.Caption = "&Opslaan"
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
Else
    editing = False
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Picture2.Cls
    updatedb (postnr)
    Command1.Enabled = True
    Command2.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    Command5.Caption = "&Bewerken"
End If
End Sub
Private Sub Command50_Click()
txtuitgeleend = Text7
Frame9.Visible = False
Frame10.Visible = True
End Sub
Private Sub Command6_Click()
If FrmFind.Visible = False Then
FrmFind.Show 0, Me
Else
Unload FrmFind
End If
End Sub
Private Sub Command60_Click()
txtrating = Text8
Frame10.Visible = False
Frame11.Visible = True
End Sub
Private Sub Command7_Click()
If editing = True Then
    tmp = MsgBox("Weet je het zeker dat je wilt gaan stoppen?" & vbCrLf & "De wijzigingen worden niet gesaved!.", vbYesNo, "M.E.")
    If tmp = 7 Then Exit Sub
End If
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Picture1.Visible = True
Picture1.Refresh
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\compact.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\compact.mdb", App.Path & "\mutlu.mdb")
Kill App.Path & "\compact.mdb"
End If
Unload FrmFind
Unload Me
End
End Sub
Private Sub Command70_Click()
txtmisc = Text9
Frame11.Visible = False
Frame12.Visible = True
End Sub
Private Sub Command8_Click()
tmpo = MsgBox("Weet je het zeker om alle films te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Call CreateNewDB("mutlu.mdb")
Form_Load
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
End Sub
Private Sub Command9_Click()
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\backup.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\backup.mdb", App.Path & "\mutlu.mdb")
End Sub
Private Sub cover_Click()
If editing = True Then
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.filename)
    If picpic.ScaleWidth >= Picture2.ScaleWidth Then picpic.Left = 0
        If picpic.ScaleHeight >= Picture2.ScaleHeight Then picpic.Top = 0
        If picpic.ScaleWidth < Picture2.ScaleWidth Then
            picpic.Left = (Picture2.ScaleWidth / 2) - (picpic.ScaleWidth / 2)
        End If
        If picpic.ScaleHeight < Picture2.ScaleHeight Then
            picpic.Top = (Picture2.ScaleHeight / 2) - (picpic.ScaleHeight / 2)
        End If
        Picture2.Cls
End If
Exit Sub
error:
MsgBox "error"
End Sub
Private Sub filmbekijken_Click()
If form1.Visible = False Then
form1.Show 0, Me
Else
Unload form1
End If
End Sub
Private Sub DataList1_Click()
 Text9 = DataList1
Dim found As Boolean, t As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.MoveFirst
For t = start To Rs.RecordCount - 1
    With Rs
        For f = 0 To .Fields.Count - 1
            If .Fields(f) <> "" Then
            test = InStr(1, Text9, .Fields(f), vbTextCompare)
                        If InStr(1, Text9, Trim(.Fields(f))) > 0 Then GoTo found
            End If
        Next f
        Rs.MoveNext
    End With
    Next t
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Exit Sub
found:
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
findstart = t
visapost t + 1
postnr = t + 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
If editing = True Then
    Picture2.Cls
    Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
    Set Rs = Db.OpenRecordset("filmlijst")
    Command1.Caption = "&Nieuwe film"
    Command5.Caption = "&Bewerken"
    If Rs.RecordCount > 0 Then
    Command2.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    End If
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
    editing = False
    visapost postnr
End If
End If
End Sub
Private Sub Form_Load()
    cds.AddItem "1 cd"
    cds.AddItem "2 cd's"
    cds.AddItem "3 cd's"
    cds.AddItem "1 dvd"
    cds.AddItem "2 dvd's"
    cds.AddItem "1 vhs"
    cds.AddItem "2 vhs's"
    
    minuten.AddItem "80 minuten"
    minuten.AddItem "81 minuten"
    minuten.AddItem "82 minuten"
    minuten.AddItem "83 minuten"
    minuten.AddItem "84 minuten"
    minuten.AddItem "85 minuten"
    minuten.AddItem "86 minuten"
    minuten.AddItem "87 minuten"
    minuten.AddItem "88 minuten"
    minuten.AddItem "89 minuten"
    minuten.AddItem "90 minuten"
    minuten.AddItem "91 minuten"
    minuten.AddItem "92 minuten"
    minuten.AddItem "93 minuten"
    minuten.AddItem "94 minuten"
    minuten.AddItem "95 minuten"
    minuten.AddItem "96 minuten"
    minuten.AddItem "97 minuten"
    minuten.AddItem "98 minuten"
    minuten.AddItem "99 minuten"
    minuten.AddItem "100 minuten"
    minuten.AddItem "101 minuten"
    minuten.AddItem "102 minuten"
    minuten.AddItem "103 minuten"
    minuten.AddItem "104 minuten"
    minuten.AddItem "105 minuten"
    minuten.AddItem "106 minuten"
    minuten.AddItem "107 minuten"
    minuten.AddItem "108 minuten"
    minuten.AddItem "109 minuten"
    minuten.AddItem "110 minuten"
    minuten.AddItem "111 minuten"
    minuten.AddItem "112 minuten"
    minuten.AddItem "113 minuten"
    minuten.AddItem "114 minuten"
    minuten.AddItem "115 minuten"
    minuten.AddItem "116 minuten"
    minuten.AddItem "117 minuten"
    minuten.AddItem "118 minuten"
    minuten.AddItem "119 minuten"
    minuten.AddItem "120 minuten"
    minuten.AddItem ">120 minuten"

    keuzemenu.AddItem "Actie"
    keuzemenu.AddItem "Animatie"
    keuzemenu.AddItem "Avontuur"
    keuzemenu.AddItem "Documentaire"
    keuzemenu.AddItem "Drama"
    keuzemenu.AddItem "Fantasie"
    keuzemenu.AddItem "Horror"
    keuzemenu.AddItem "MANGA"
    keuzemenu.AddItem "Praatfilm"
    keuzemenu.AddItem "Romantisch"
    keuzemenu.AddItem "Sex"
    keuzemenu.AddItem "Science-fiction"
    keuzemenu.AddItem "Tekenfilm"
    keuzemenu.AddItem "Thriller"
        
    Jaar.AddItem "2003"
    Jaar.AddItem "2002"
    Jaar.AddItem "2001"
    Jaar.AddItem "2000"
    Jaar.AddItem "1999"
    Jaar.AddItem "1998"
    Jaar.AddItem "1997"
    Jaar.AddItem "1996"
    Jaar.AddItem "1995"
    Jaar.AddItem "1994"
    Jaar.AddItem "1993"
    Jaar.AddItem "1992"
    Jaar.AddItem "1991"
    Jaar.AddItem "1990"
    Jaar.AddItem "1989"
    Jaar.AddItem "1988"
    Jaar.AddItem "1987"
    Jaar.AddItem "1986"
    Jaar.AddItem "1985"
    Jaar.AddItem "1984"
    Jaar.AddItem "1983"
    Jaar.AddItem "1982"
    Jaar.AddItem "1981"

postnr = 1
If Dir(App.Path & "\mutlu.mdb") <> "" Then GoTo continue
CreateNewDB "mutlu.mdb"
continue:
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
    visapost (1)
    Open App.Path & "\mutlu.mdb" For Binary As #1
    g = LOF(1)
    Close #1
Else
    StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
    Command8.Enabled = False
    txtpost = "0"
    Command2.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
End If
Open App.Path & "\mutlu.mdb" For Binary As #1
g = LOF(1)
Close #1
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
End Sub
Public Sub savetodb()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
    Set Rs = Db.OpenRecordset("filmlijst")
    On Error Resume Next
        Rs.AddNew
        Rs.Fields(0) = Trim(txtfilmnaam)
        Rs.Fields(1) = Trim(txtacteurs)
        Rs.Fields(2) = Trim(txtgenre)
        Rs.Fields(3) = Trim(txttype)
        Rs.Fields(4) = Trim(txtspeelduur)
        Rs.Fields(5) = Trim(txtjaar)
        Rs.Fields(6) = Trim(txtrating)
        Rs.Fields(7) = Trim(txtuitgeleend)
        Rs.Fields(8) = Trim(txtmisc)
        Rs.Fields(9) = Trim(Now)
        Rs.Fields(0) = Trim(txtcd)
        Rs.Fields(11) = Trim(aantalcds)
        Rs.Fields(12) = Trim(ondertiteling)
        If picpic.Picture > "" Then
            Call SavePicture(picpic.Picture, App.Path & "\tmpfile")
            Dim strFromFile As String
            Dim lngFileSize As Long
            Dim FileNum As Integer
            FileNum = FreeFile
            lngFileSize = FileLen(App.Path & "\tmpfile")
            strFromFile = String(lngFileSize, " ")
            Open App.Path & "\tmpfile" For Binary As FileNum
            Get FileNum, , strFromFile
            Close FileNum
            Rs.Fields(10) = strFromFile
            Kill App.Path & "\tmpfile"
        End If
        Rs.Update
        Rs.MoveLast
        lblcreated.Caption = "Film toegevoegd " & Trim(Rs.Fields(9))
        StatusBar1.Panels(1).Text = "Er zijn" & Str(Rs.RecordCount) & " films in de database. Film nr. " & Str(Rs.RecordCount) & "."
        txtpost.Text = Str(Rs.RecordCount)
        postnr = Rs.RecordCount
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        Open App.Path & "\mutlu.mdb" For Binary As #1
        g = LOF(1)
        Close #1
        StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
        visapost (postnr)
        txtfilmnaam.Locked = True
        txtacteurs.Locked = True
        txtgenre.Locked = True
        txttype.Locked = True
        txtspeelduur.Locked = True
        txtjaar.Locked = True
        txtrating.Locked = True
        txtuitgeleend.Locked = True
        txtmisc.Locked = True
        txtcd.Locked = True
        aantalcds.Locked = True
        ondertiteling.Locked = True
End Sub
Private Sub GeavanceerdeFilmInfo_Click()
If filminfo.Visible = False Then
filminfo.Show 0, Me
Else
Unload filminfo
End If
End Sub
Private Sub info_Click()
If Form2.Visible = False Then
Form2.Show 0, Me
Else
Unload Form2
End If
End Sub
Private Sub kazaa_Click()
If Form3.Visible = False Then
Form3.Show 0, Me
Else
Unload Form3
End If
End Sub
Private Sub maakbackup_Click()
Kill App.Path & "\backup.mdb"
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\backup.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\backup.mdb", App.Path & "\mutlu.mdb")
End
End Sub
Private Sub mmdbnetwerk_Click()
If frmHaupt.Visible = False Then
frmHaupt.Show 0, Me
Else
Unload frmHaupt
End If
End Sub
Private Sub Nieuwe_film_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    editing = True
    Command2.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.SetFocus
    Command1.Caption = "&Opslaan"
Else
    Picture2.Cls
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Command2.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    savetodb
    editing = False
    Command1.Caption = "&Nieuwe film"
End If
End Sub
Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image11.Visible = False
Image10.Visible = False
Image9.Visible = False
Image8.Visible = False
Image7.Visible = False
Image6.Visible = False
Image5.Visible = False
Image4.Visible = False
Image3.Visible = False
Image2.Visible = False
Image1.Visible = True
End Sub
Private Sub Image10_Click()
Command36.Visible = True
Command28.Visible = True
Command24.Visible = True
Command26.Visible = True
Command25.Visible = True
Command27.Visible = True
Command20.Visible = True
Command17.Visible = True
Command15.Visible = True
Command13.Visible = True
Command16.Visible = True
Command37.Visible = True
Command38.Visible = True
If editing = False Then
    Picture11.Enabled = False
    Picture12.Enabled = False
    Picture4.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    End If
    editing = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    Command5.Caption = "&Opslaan"
    Picture5.Visible = False
    Picture15.Visible = True
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
Else
    editing = False
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Picture2.Cls
    updatedb (postnr)
    Command1.Enabled = True
    Command2.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    Command5.Caption = "&Bewerken"
End If
End Sub
Private Sub Image11_Click()
Command37.Visible = False
Command36.Visible = False
Command23.Visible = False
Command14.Visible = False
Command28.Visible = False
Command30.Visible = False
Command24.Visible = False
Command26.Visible = False
Command25.Visible = False
Command27.Visible = False
Command20.Visible = False
Command17.Visible = False
Command29.Visible = False
Command15.Visible = False
Command13.Visible = False
Command16.Visible = False
Command38.Visible = False
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    End If
    editing = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    Command5.Caption = "&Opslaan"
    Picture5.Visible = True
    Picture15.Visible = False
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
Else
    editing = False
    Picture11.Enabled = True
    Picture12.Enabled = True
    Picture4.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Picture2.Cls
    updatedb (postnr)
    Command1.Enabled = True
    Command2.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    Command5.Caption = "&Bewerken"
    Picture15.Visible = False
    Picture5.Visible = True
End If
End Sub
Private Sub Image12_Click()
Image12.Picture = LoadPicture(App.Path & "\lens.jpg")
End Sub
Private Sub Image13_Click()
Image13.Picture = LoadPicture(App.Path & "\banderas.jpg")
End Sub

Private Sub Image14_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
End Sub

Private Sub Image15_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 1
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost postnr
End If
End Sub

Private Sub Image2_Click()
If editing = True Then
    tmp = MsgBox("Weet je het zeker dat je wilt gaan stoppen?" & vbCrLf & "De wijzigingen worden niet gesaved!.", vbYesNo, "M.E.")
    If tmp = 7 Then Exit Sub
End If
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Picture1.Visible = True
Picture1.Refresh
End If
Unload FrmFind
Unload Me
End
End Sub
Private Sub Image3_Click()
Dim f As New FrmFind
           f.Show
End Sub
Private Sub Image4_Click()
tmpo = MsgBox("Weet je het zeker om alle films te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Call CreateNewDB("mutlu.mdb")
Form_Load
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
End Sub
Private Sub Image5_Click()
tmpo = MsgBox("Weet je het zeker om de geselecteerde film te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Dim tmp As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.Move postnr - 1
Rs.Delete
postnr = Rs.RecordCount
txtpost = postnr
If Rs.RecordCount > 0 Then
tmp = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost tmp
Open App.Path & "\mutlu.mdb" For Binary As #1
g = LOF(1)
Close #1
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Else
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtrating.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
    txtcd.Locked = True
    aantalcds.Locked = True
    ondertiteling.Locked = True
StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
Command8.Enabled = False
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End If
End Sub
Private Sub Image6_Click()
Dim f As New form1
           f.Show
End Sub
Private Sub Image7_Click()
If editing = True Then
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.filename)
    If picpic.ScaleWidth >= Picture2.ScaleWidth Then picpic.Left = 0
        If picpic.ScaleHeight >= Picture2.ScaleHeight Then picpic.Top = 0
        If picpic.ScaleWidth < Picture2.ScaleWidth Then
            picpic.Left = (Picture2.ScaleWidth / 2) - (picpic.ScaleWidth / 2)
        End If
        If picpic.ScaleHeight < Picture2.ScaleHeight Then
            picpic.Top = (Picture2.ScaleHeight / 2) - (picpic.ScaleHeight / 2)
        End If
        Picture2.Cls
End If
Exit Sub
error:
MsgBox "error"
End Sub
Private Sub Image8_Click()
If editing = False Then
    Picture11.Enabled = False
    Picture12.Enabled = False
    Picture5.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    editing = True
    Command2.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.SetFocus
    Command1.Caption = "&Opslaan"
    Picture3.Visible = True
    Picture4.Visible = False
Else
    Picture2.Cls
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Command2.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    savetodb
    editing = False
    Command1.Caption = "&Nieuwe film"
End If
Command36.Visible = True
Command28.Visible = True
Command24.Visible = True
Command26.Visible = True
Command25.Visible = True
Command27.Visible = True
Command20.Visible = True
Command17.Visible = True
Command15.Visible = True
Command13.Visible = True
Command16.Visible = True
Command37.Visible = True
Command38.Visible = True
txttype = Text1
txtspeelduur = Text3
txtjaar = Text2
txtgenre = Text4
txtuitgeleend = Text5
txtrating = Text6
txtmisc = Text7
End Sub
Private Sub Image9_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    editing = True
    Command2.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command8.Enabled = False
    txtfilmnaam.Locked = False
    txtacteurs.Locked = False
    txtgenre.Locked = False
    txttype.Locked = False
    txtspeelduur.Locked = False
    txtjaar.Locked = False
    txtrating.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtcd.Locked = False
    aantalcds.Locked = False
    ondertiteling.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.SetFocus
    Command1.Caption = "&Opslaan"
    Picture3.Visible = True
    Picture4.Visible = False
Else
 Picture11.Enabled = True
    Picture12.Enabled = True
    Picture5.Enabled = True
    Picture2.Cls
    Command3.Enabled = True
    Command4.Enabled = True
    txtpost.Locked = False
    Command2.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command8.Enabled = True
    savetodb
    editing = False
    Command1.Caption = "&Nieuwe film"
    Picture3.Visible = False
    Picture4.Visible = True
End If
Command37.Visible = False
Command36.Visible = False
Command23.Visible = False
Command14.Visible = False
Command28.Visible = False
Command30.Visible = False
Command24.Visible = False
Command26.Visible = False
Command25.Visible = False
Command27.Visible = False
Command20.Visible = False
Command17.Visible = False
Command29.Visible = False
Command15.Visible = False
Command13.Visible = False
Command16.Visible = False
Command38.Visible = False
End Sub
Private Sub Jaar_Change()
txtjaar = Jaar
End Sub
Private Sub keuzemenu_Change()
txtgenre = keuzemenu
End Sub
Private Sub Option1_Click()
Textkeuze = "DivX"
End Sub
Private Sub Option2_Click()
Textkeuze = "DVD"
End Sub
Private Sub Option3_Click()
Textkeuze = "VHS"
End Sub
Private Sub Option4_Click()
Textkeuze = "VCD"
End Sub
Private Sub Option5_Click()
Textkeuze = "SVCD"
End Sub
Private Sub Option6_Click()
Textkeuze = "PDivX"
End Sub
Private Sub Option7_Click()
ondertiteling = "Ja"
End Sub
Private Sub Option8_Click()
ondertiteling = "Nee"
End Sub
Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image14.Visible = True


End Sub
Private Sub Picture12_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image15.Visible = True


End Sub
Private Sub Picture13_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image7.Visible = True
End Sub

Private Sub picture14_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image15.Visible = False
Image14.Visible = False
Image11.Visible = False
Image10.Visible = False
Image9.Visible = False
Image8.Visible = False
Image7.Visible = False
Image6.Visible = False
Image5.Visible = False
Image4.Visible = False
Image3.Visible = False
Image2.Visible = False
End Sub

Private Sub Picture15_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image11.Visible = True
End Sub
Private Sub Picture16_Click()
    Dim frmB As New frmBrowser
    frmB.StartingAddress = browserstring + "http://us.imdb.com/Tsearch?title=" + txtcd.Text + "&restrict=Movies"
    frmB.Show
End Sub
Private Sub Picture17_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 10000
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
End Sub
Private Sub Picture18_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 10000
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost postnr
End If
End Sub
Private Sub Picture2_DblClick()
If editing = True Then
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.filename)
    If picpic.ScaleWidth >= Picture2.ScaleWidth Then picpic.Left = 0
        If picpic.ScaleHeight >= Picture2.ScaleHeight Then picpic.Top = 0
        If picpic.ScaleWidth < Picture2.ScaleWidth Then
            picpic.Left = (Picture2.ScaleWidth / 2) - (picpic.ScaleWidth / 2)
        End If
        If picpic.ScaleHeight < Picture2.ScaleHeight Then
            picpic.Top = (Picture2.ScaleHeight / 2) - (picpic.ScaleHeight / 2)
        End If
        Picture2.Cls
End If
Exit Sub
error:
MsgBox "error"
End Sub
Private Sub Picture2_KeyPress(KeyAscii As Integer)
If editing = True Then
If KeyAscii Then
    If Clipboard.GetFormat(2) = True Then
        picpic.Visible = True
        picpic.Picture = Clipboard.GetData(2)
        If picpic.ScaleWidth > Picture2.ScaleWidth Then picpic.Left = 0
        If picpic.ScaleHeight > Picture2.ScaleHeight Then picpic.Top = 0
        If picpic.ScaleWidth < Picture2.ScaleWidth Then
            picpic.Left = (Picture2.ScaleWidth / 2) - (picpic.ScaleWidth / 2)
        End If
        If picpic.ScaleHeight < Picture2.ScaleHeight Then
            picpic.Top = (Picture2.ScaleHeight / 2) - (picpic.ScaleHeight / 2)
        End If
        Picture2.Cls
    End If
End If
End If
End Sub
Private Sub picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image9.Visible = True
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image8.Visible = True
End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image10.Visible = True
End Sub
Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image5.Visible = True
End Sub
Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image4.Visible = True
End Sub
Private Sub Picture8_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image6.Visible = True
End Sub
Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image3.Visible = True
End Sub
Private Sub starten_Click()
If frmHaupt.Visible = False Then
frmHaupt.Show 0, Me
Else
Unload frmHaupt
End If
End Sub
Private Sub print_Click()
Printer.Print DataList1.Text
Printer.EndDoc
End Sub
Private Sub Printfilminfo_Click()
Printer.Print txtmisc.Text
Printer.EndDoc
End Sub
Private Sub printscreen_Click()
End Sub
Private Sub registreerde_Click()
If Formreg3.Visible = False Then
Formreg3.Show 0, Me
Else
Unload Formreg3
End If
End Sub
Private Sub ratingbalk_Change()
txtrating = ratingbalk
End Sub
Private Sub Textkeuze_Change()
txttype = Textkeuze
End Sub

Private Sub txtpost_KeyDown(KeyCode As Integer, Shift As Integer)
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
If KeyCode = 13 Then
    If txtpost < 1 Then txtpost = 1
    If txtpost > Rs.RecordCount Then txtpost = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
post = txtpost
visapost txtpost
End If
Else
txtpost = "0"
End If
End Sub
Public Sub updatedb(post As Integer)
    Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
    Set Rs = Db.OpenRecordset("filmlijst")
    On Error Resume Next
        Rs.Move post - 1
        Rs.Edit
        Rs.Fields(0) = Trim(txtfilmnaam)
        Rs.Fields(1) = Trim(txtacteurs)
        Rs.Fields(2) = Trim(txtgenre)
        Rs.Fields(3) = Trim(txttype)
        Rs.Fields(4) = Trim(txtspeelduur)
        Rs.Fields(5) = Trim(txtjaar)
        Rs.Fields(6) = Trim(txtrating)
        Rs.Fields(7) = Trim(txtuitgeleend)
        Rs.Fields(8) = Trim(txtmisc)
        Rs.Fields(9) = Trim(Now)
        Rs.Fields(0) = Trim(txtcd)
        Rs.Fields(11) = Trim(aantalcds)
        Rs.Fields(12) = Trim(ondertiteling)
        If picpic.Picture > "" Then
            Call SavePicture(picpic.Picture, App.Path & "\tmpfile")
            Dim strFromFile As String
            Dim lngFileSize As Long
            Dim FileNum As Integer
            FileNum = FreeFile
            lngFileSize = FileLen(App.Path & "\tmpfile")
            strFromFile = String(lngFileSize, " ")
            Open App.Path & "\tmpfile" For Binary As FileNum
            Get FileNum, , strFromFile
            Close FileNum
            Rs.Fields(10) = strFromFile
            Kill App.Path & "\tmpfile"
        End If
        Rs.Update
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        txtfilmnaam.Locked = True
        txtacteurs.Locked = True
        txtgenre.Locked = True
        txttype.Locked = True
        txtspeelduur.Locked = True
        txtjaar.Locked = True
        txtrating.Locked = True
        txtuitgeleend.Locked = True
        txtmisc.Locked = True
        txtcd.Locked = True
        aantalcds.Locked = True
        ondertiteling.Locked = True
        Open App.Path & "\mutlu.mdb" For Binary As #1
        g = LOF(1)
        Close #1
        StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
        visapost post
End Sub
Private Sub uitgeleendaan_Click()
If form1.Visible = False Then
Form4.Show 0, Me
Else
Unload Form4
End If
End Sub
Private Sub Verwijder_film_Click()
tmpo = MsgBox("Weet je het zeker om de geselecteerde film te verwijderen?", vbYesNo, "M.E.")
If tmpo = 7 Then Exit Sub
Dim tmp As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.Move postnr - 1
Rs.Delete
postnr = Rs.RecordCount
txtpost = postnr
If Rs.RecordCount > 0 Then
tmp = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
visapost tmp
Open App.Path & "\mutlu.mdb" For Binary As #1
g = LOF(1)
Close #1
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Else
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtrating = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtcd = ""
    aantalcds = ""
    ondertiteling = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtrating.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
    txtcd.Locked = True
    aantalcds.Locked = True
    ondertiteling.Locked = True
StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
Command8.Enabled = False
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End If
End Sub
Private Sub Volgendefilm_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 1
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
'txtpost = Str(postnr)
visapost postnr
End If
End Sub
Private Sub Vorigefilm_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
End Sub
Private Sub Zoek_film_Click()
If FrmFind.Visible = False Then
FrmFind.Show 0, Me
Else
Unload FrmFind
End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image2.Visible = True
End Sub
Private Sub zoekfilmapart_Click()
If frmList.Visible = False Then
frmList.Show 0, Me
Else
Unload frmList
End If
End Sub
Private Sub Check1_Click()
ratingbalk = "*"
End Sub
Private Sub Check10_Click()
ratingbalk = "**********"
End Sub
Private Sub Check2_Click()
ratingbalk = "**"
End Sub
Private Sub Check3_Click()
ratingbalk = "***"
End Sub
Private Sub Check4_Click()
ratingbalk = "****"
End Sub
Private Sub Check5_Click()
ratingbalk = "*****"
End Sub
Private Sub Check6_Click()
ratingbalk = "******"
End Sub
Private Sub Check7_Click()
ratingbalk = "*******"
End Sub
Private Sub Check8_Click()
ratingbalk = "********"
End Sub
Private Sub Check9_Click()
ratingbalk = "*********"
End Sub
