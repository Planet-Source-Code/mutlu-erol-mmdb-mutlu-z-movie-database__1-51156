VERSION 5.00
Begin VB.Form help 
   Caption         =   "help"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4290
   Icon            =   "help.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5115
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command14 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   62
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   61
         Top             =   4200
         Width           =   2055
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   480
         Picture         =   "help.frx":09CA
         ScaleHeight     =   2985
         ScaleWidth      =   1545
         TabIndex        =   60
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Acteurs"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label35 
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 2"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label36 
         BackColor       =   &H00000000&
         Caption         =   $"help.frx":20CD
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   2520
         TabIndex        =   63
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   55
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Command15 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   54
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Speelduur"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   58
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label45 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie        ""STAP 3"""
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
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command20 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000008&
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 4"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command23 
         Caption         =   "Selecteer cover"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   44
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command26 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "10"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie     ""STAP 10"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command31 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Jaar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label27 
         BackColor       =   &H00000000&
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 5"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command40 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Genre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000000&
         Caption         =   "6"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 6"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command27 
         Caption         =   "Opslaan en eindig"
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000000&
         Caption         =   "11"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie     ""Stap 11"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command50 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Uitgeleend "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000000&
         Caption         =   "7"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 7"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command60 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "Rating"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label30 
         BackColor       =   &H00000000&
         Caption         =   "8"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label40 
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 8"""
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3840
         Width           =   3015
      End
      Begin VB.CommandButton Command70 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "Film info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label31 
         BackColor       =   &H00000000&
         Caption         =   "9"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label39 
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie        ""STAP 9"""
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
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   360
         Picture         =   "help.frx":21E1
         ScaleHeight     =   2745
         ScaleWidth      =   1665
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command13 
         Caption         =   "volgende"
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Naam"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "MMDB Help-functie       ""STAP 1"""
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
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000012&
         Caption         =   "1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label34 
         BackColor       =   &H00000000&
         Caption         =   $"help.frx":38E4
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
