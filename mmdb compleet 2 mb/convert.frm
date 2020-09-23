VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Convert 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " HTML / ASP Convertor"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "convert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10020
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   240
   End
   Begin VB.CommandButton cmdhelp 
      BackColor       =   &H00C0C000&
      Height          =   495
      Left            =   7680
      Picture         =   "convert.frx":09CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Picture         =   "convert.frx":0E0C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Picture         =   "convert.frx":124E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Picture         =   "convert.frx":1690
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin MSComDlg.CommonDialog help 
         Left            =   1440
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   $"convert.frx":1AD2
         ForeColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame step1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   35
         ImageHeight     =   34
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "convert.frx":1B99
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "convert.frx":2A43
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cbodatabase 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList1"
      End
      Begin VB.Frame frmaccess 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3855
         Begin VB.CommandButton cmdbrowse 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Browse..."
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1800
            Width           =   1455
         End
         Begin MSComDlg.CommonDialog cd 
            Left            =   3000
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtdatabasepath 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "C:\Program Files\Mutlu'z Movie database\mutlu.mdb"
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C:\Program Files\Mutlu'z Movie database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   3345
         End
      End
      Begin VB.Frame frmsql 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Authentication"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3855
         Begin VB.CommandButton cmdconnect 
            BackColor       =   &H0000C000&
            Caption         =   "Connect"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox cbosqldatabase 
            BackColor       =   &H0000C0C0&
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtpass 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   18
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtuser 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtserver 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database Name :"
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3960
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SQL Server"
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User name"
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   765
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Selecteer de database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   2145
      End
   End
   Begin VB.Frame step5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   2520
      TabIndex        =   44
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdfinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voltooien"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3000
         Width           =   1935
      End
   End
   Begin VB.Frame step4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   2520
      TabIndex        =   38
      Top             =   0
      Width           =   4335
      Begin RichTextLib.RichTextBox Finalresult 
         Height          =   3135
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"convert.frx":39C5
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3135
         Left            =   120
         Top             =   840
         Width           =   4095
      End
      Begin VB.Image Image7 
         Height          =   360
         Left            =   360
         Picture         =   "convert.frx":3A47
         Stretch         =   -1  'True
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FF0000&
         Height          =   135
         Left            =   2640
         TabIndex        =   50
         Top             =   4065
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "= Tables  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   49
         Top             =   4020
         Width           =   870
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "= Fields"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   48
         Top             =   4020
         Width           =   675
      End
      Begin VB.Label Label16 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   960
         TabIndex        =   39
         Top             =   4065
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Is deze informatie juist ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   47
         Top             =   480
         Width           =   2340
      End
   End
   Begin VB.Frame step3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   2520
      TabIndex        =   33
      Top             =   0
      Width           =   4335
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   52
         Top             =   3960
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   480
         TabIndex        =   51
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtpath 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "C:\Program Files\Mutlu'z Movie database"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton opthtml 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Data In HTML "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optasp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Data In ASP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1440
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Nieuw ASP converter!! Maar is in ontwikkelfase"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   960
         Width           =   3495
      End
      Begin VB.Image Image4 
         Height          =   555
         Left            =   2520
         Picture         =   "convert.frx":4189
         Top             =   360
         Width           =   435
      End
      Begin VB.Image Image3 
         Height          =   540
         Left            =   2520
         Picture         =   "convert.frx":4E83
         Top             =   1200
         Width           =   510
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   120
         Picture         =   "convert.frx":5D65
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label heading 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Selecteer de path van MUTLU.MDB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   480
         TabIndex        =   37
         Top             =   1920
         Width           =   3075
      End
   End
   Begin VB.Frame step2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   2520
      TabIndex        =   25
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.ListView lstfields 
         Height          =   1455
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fields"
            Object.Width           =   4498
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2141
         EndProperty
      End
      Begin MSComctlLib.ListView lstaddedfields 
         Height          =   1095
         Left            =   120
         TabIndex        =   41
         Top             =   3240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Table"
            Object.Width           =   2540
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Field"
            Object.Width           =   4523
            ImageIndex      =   2
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2040
         Top             =   2880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "convert.frx":61A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "convert.frx":65F9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton optdeselect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Deselecteer All"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   32
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optselect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Selecteer All"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H0000C000&
         Caption         =   "Add Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox cbotable 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toegevoegde velden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added Fields"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   105
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voeg veld toe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select The Table/Query Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   26
         Top             =   360
         Width           =   2610
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   7800
      TabIndex        =   11
      Top             =   5760
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Annuleer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6000
      TabIndex        =   9
      Top             =   5040
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volgende"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   8
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   5040
      Width           =   510
   End
End
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbodatabase_Click()

  If Me.cbodatabase.SelectedItem.Text = "MS Access" Then
     Me.frmaccess.ZOrder
     Me.cmdbrowse.TabIndex = 0
  Else
     Me.frmsql.ZOrder
     Me.txtdatabasepath.Text = ""
     Me.txtserver.Text = "(local)"
     Me.txtpass.Text = ""
     Me.txtuser.Text = ""
     Me.txtserver.SelStart = 0
     Me.txtserver.SelLength = 7
     Me.txtserver.SetFocus
  End If

  Me.cbotable.Clear
  Me.lstaddedfields.ListItems.Clear
  
End Sub

Private Sub cbotable_Click()
On Error Resume Next

  If Field.State = 1 Then Field.Close
  Field.Open "select * from [" & Me.cbotable.Text & "]", cn, adOpenDynamic, adLockOptimistic
  
  Me.lstfields.ListItems.Clear

  For i = 0 To Field.Fields.Count - 1
     Set X = Me.lstfields.ListItems.add(, , Field(i).Name)
     X.SubItems(1) = "-"
     X.ForeColor = vbRed
  Next
  
  Addmarker
  Temp_Table = Me.cbotable.Text

End Sub

Public Sub Addmarker()
On Error GoTo jump

    '-----------------------------------
    'To indicate the user that these
    'fields are already added
    '-----------------------------------
    With Me.lstaddedfields
    
      For i = 1 To Me.lstfields.ListItems.Count
        For j = 1 To .ListItems.Count
          Set X = .ListItems.Item(j)
          If X.SubItems(1) = Me.lstfields.ListItems.Item(i) And .ListItems.Item(j) = Me.cbotable Then
            Set X = Me.lstfields.ListItems.Item(i)
            X.SubItems(1) = "Added"
          End If
        Next
      Next
      
    End With

Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Private Sub cmdadd_Click()
On Error GoTo jump

    '----------------------------------------------------------
    'THIS LOOP CHECK WHETHER SELECTED FIELDS ARE ALREADY ADDED
    'IN THE LIST OR NOT IF NOT THEN FIELD WILL ADDED
    '----------------------------------------------------------
    For i = 1 To Me.lstfields.ListItems.Count
      Set X = Me.lstfields.ListItems.Item(i)
      If Me.lstfields.ListItems.Item(i).Checked = True Then
        If X.SubItems(1) <> "Added" Then
          Set X = lstaddedfields.ListItems.add(, , cbotable.Text)
          X.SubItems(1) = Me.lstfields.ListItems.Item(i)
        Else
          Field_added = Field_added & Me.lstfields.ListItems.Item(i) & vbCrLf
        End If
      End If
    Next

   'SET THE STATUS OF FIELD/MARKER
   Addmarker
   
   If Field_added <> "" Then MsgBox "Following fields are already added" & vbCrLf _
                                  + "---------------------------------------" & vbCrLf _
                                  + Field_added, vbExclamation
   Field_added = ""
 
   optdeselect_Click
   Me.optselect.Value = False
   Me.optdeselect.Value = False
 
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Private Sub cmdbrowse_Click()
On Error GoTo jump

    With cd
     .DialogTitle = "Select The Database"
     .CancelError = False
     .Filter = "Access File (*.Mdb)|*.mdb"
     .ShowOpen
      txtdatabasepath.Text = .filename
    End With
    
    If Me.txtdatabasepath.Text <> "" Then
      If cn.State = 1 Then cn.Close
      cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & Trim(Me.txtdatabasepath.Text)
      cn.CursorLocation = adUseClient
      Proceed = True
    End If
    
Exit Sub
jump:
MsgBox Err.Description, vbCritical
Proceed = False
End Sub

Private Sub cmdcancel_Click()
Unload Me
'If MsgBox("Exit From Setup", vbYesNo + vbQuestion) = vbYes Then
 ' MsgBox "Setup Not Completed Yet", vbInformation
  'End
'End If
End Sub

Private Sub cmdconnect_Click()
On Error GoTo jump

    SqlConnect
    MsgBox "SQL Server Successfully Connected", vbInformation
    
    Screen.MousePointer = vbHourglass
    
    If SQL_Database.State = 1 Then SQL_Database.Close
    SQL_Database.Open "sp_helpdb", cn, adOpenDynamic, adLockOptimistic
    
    Me.cbosqldatabase.Clear
    
    While Not SQL_Database.EOF
      If SQL_Database.Fields("name") <> "master" And SQL_Database.Fields("name") <> "model" And _
        SQL_Database.Fields("name") <> "msdb" Then
        Me.cbosqldatabase.AddItem SQL_Database.Fields("name")
      End If
    SQL_Database.MoveNext
    Wend
    
    Screen.MousePointer = vbArrow
    Proceed = True
Exit Sub
jump:
MsgBox Err.Description, vbCritical
Proceed = False
End Sub

Private Sub cmdfinish_Click()

   If Me.optasp.Value = True Then 'ASP
     
    ASPConverter 'all tables will be converted in ASP files
     
     '----------------------------------------------------------
     ' NOTE : DO NOT DELETE THIS LINE
     '        THIS LINE WILL COPY ADOVBS.INC FILE IN A FOLDER
     '        TO RUN ASP FILE
     '----------------------------------------------------------
   Set fs = CreateObject("scripting.FileSystemObject")
   fs.CopyFile App.Path & "\Adovbs\Adovbs.inc", FOLDER_PATH & "\"
     
     
     MsgBox "Alle ASP files zijn Successfull gemaakt" & vbCrLf & vbCrLf _
           + "Path : " & FOLDER_PATH & vbCrLf & vbCrLf _
           
           ' + " NOTE : Always Run MAIN.HTML Under This Folder As A Main Menu", vbInformation
            
   
   Else 'Html
     
     HTMLConverter 'all tables will be converted in HTML files
     
     MsgBox "Alle Html files zijn Successfull gemaakt" & vbCrLf _
           + "Path : " & FOLDER_PATH & vbCrLf & vbCrLf _
         
         '  + " NOTE : Always Run MAIN.HTML Under This Folder As A Main Menu", vbInformation
   End If
   
   'MsgBox "For Any Suggestion Or Query Write to deepakmailto@rediffmail.com", vbInformation
   'shel = ShellExecute(o&, vbNullString, FOLDER_PATH & "/main.html", vbNullString, vbNullString, SW_Shownormal)
   
 '  End
   
End Sub

Private Sub cmdhelp_Click()

cd.HelpFile = App.Path & "\ONLINE HELP.hlp"
cd.HelpCommand = cdlHelpContents
cd.ShowHelp

End Sub

Private Sub cmdnext_Click()

    If Me.cbodatabase.SelectedItem.Text = "MS Access" Then
      If Me.txtdatabasepath.Text = "" Then
        MsgBox "Select The Database Path", vbExclamation
        Exit Sub
      End If
     
      'store the database information
      Database_path = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.txtdatabasepath.Text & ";Persist Security Info=False"
      AccessTables
      
      If Check_Database = "" Then
       Check_Database = Me.txtdatabasepath.Text
      ElseIf Check_Database <> Me.txtdatabasepath.Text Then
        Me.lstaddedfields.ListItems.Clear
        Me.lstfields.ListItems.Clear
        Me.txtpath.Text = ""
      End If
      
      
    
    ElseIf Me.cbodatabase.SelectedItem.Text = "SQL Server" Then
      If Me.cbosqldatabase.Text = "" Then
         MsgBox "Select The Database Name", vbExclamation
         Exit Sub
      End If
      
      SqlConnect
      'store the database information
      Database_path = "provider=sqloledb;server=" & Trim(Convert.txtserver) & ";user id=" & Trim(Convert.txtuser) & ";password=" & Trim(Convert.txtpass) & ";database=" & Convert.cbosqldatabase.Text
      SQLTables
    End If
    
 
    If Proceed = True Then
      wiz_counter = wiz_counter + 1
      Select Case wiz_counter
      Case 1
        If Me.cbodatabase.SelectedItem.Text = "MS Access" Then Check_Database = Me.txtdatabasepath.Text
        step2.ZOrder
        Me.cmdprev.Enabled = True
      Case 2
        If Me.lstaddedfields.ListItems.Count < 1 Then
           MsgBox "No Field Is Added In Added Field List", vbInformation
           Me.cbotable.Text = Temp_Table
           wiz_counter = 1
           Exit Sub
        End If
        step3.ZOrder
      Case 3
       If Me.txtpath.Text = "" Then
         MsgBox heading.Caption, vbInformation
         wiz_counter = 2
         Exit Sub
       End If
       Rem ---------------------------------------------------
       Rem CHECK THAT IF CONVERTER IS ASP THEN USER HAVE
       Rem TO SELECT THE VIRTUAL DIRECTORY NOT ONLY THE DRIVE
       Rem ---------------------------------------------------
        If Me.optasp.Value = True Then
          If Mid(Me.txtpath.Text, InStr(1, Me.txtpath.Text, "\") + 2) = "" Then
            MsgBox "You Do Not Specify Your Virtual Directory", vbExclamation
            wiz_counter = 2
            Exit Sub
          End If
        End If
        
        Final
        step4.ZOrder
      Case 4
        step5.ZOrder
        Me.cmdnext.Enabled = False
      End Select
    End If
    
End Sub

Public Sub Final()
On Error GoTo jump

  DistinctName
 
  With Me.Finalresult
    
    .Text = ""
    For i = 1 To UBound(arr()) - 1
     .SelColor = vbRed
     .SelBold = True
     .SelUnderline = True
     .SelText = .SelText & vbCrLf
     .SelText = arr(i) & vbCrLf
     .SelBold = False
      For j = 1 To Me.lstaddedfields.ListItems.Count
        Set X = Me.lstaddedfields.ListItems.Item(j)
        If arr(i) = Me.lstaddedfields.ListItems.Item(j) Then
          .SelBullet = True
          .SelColor = vbBlue
          .SelText = X.SubItems(1) & vbCrLf
          .SelBullet = False
          .SelColor = vbBlack
        End If
      Next
    Next
    
  End With
  Me.Finalresult.SelStart = 0
  Erase arr 'CLEAR THE ARRAY CONTAINTS
  ReDim arr(1)

Exit Sub
jump:
MsgBox Err.Description, vbCritical
End Sub

Public Sub DistinctName()
On Error GoTo jump

    For j = 1 To Me.lstaddedfields.ListItems.Count
      For i = 0 To UBound(arr()) - 1
        If Me.lstaddedfields.ListItems.Item(j) <> arr(i) Then
          Table_Added = True
        Else
          Table_Added = False
        End If
      Next
      
      If Table_Added = True Then
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr) - 1) = Me.lstaddedfields.ListItems.Item(j)
      End If
    Next

Exit Sub
jump:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdprev_Click()
On Error Resume Next

  If Proceed = True Then
    wiz_counter = wiz_counter - 1
    Select Case wiz_counter
    Case 0
      step1.ZOrder
      Me.cmdprev.Enabled = False
    Case 1
      step2.ZOrder
      Me.cbotable.Text = Temp_Table
    Case 2
      step3.ZOrder
    Case 3
      step4.ZOrder
      Me.cmdnext.Enabled = True
    End Select
  End If
  
End Sub

Private Sub Dir1_Change()
txtpath.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    '------------------------------------------------------------------------------
    'note : I have fix the CD drive as E: but it ca be chage on different machines
    '       i am not able to get the cd drive so please adjust it
    '------------------------------------------------------------------------------

    If (Drive1.Drive = "A:" Or Drive1.Drive = "a:") Or (Drive1.Drive = "E:" Or Drive1.Drive = "e:") Then
       MsgBox "Sorry This Drive Not Allowed"
    Else
       Dir1.Path = Drive1.Drive
    End If

End Sub

Private Sub Form_Load()
  Initlize
End Sub

Public Sub Initlize()

   Top = 1890
   Left = 2355
   Height = 5685
   Width = 7170
   wiz_counter = 0
   ReDim arr(1)
   step1.ZOrder

   With Me.cbodatabase
    .ComboItems.add , "access", "MS Access", 1
    .ComboItems.add , "sql", "SQL Server", 2
   End With

   Me.cbodatabase.ComboItems("access").Selected = True
   cbodatabase_Click
   
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image5_Click()

End Sub

Private Sub lstaddedfields_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo jump

    If KeyCode = vbKeyDelete Then
      '------------------------------------------------------------
      'THIS LINE WILL DELETE THE FIELD FROM THE LIST AND RESTORE IT
      '------------------------------------------------------------
      With Me.lstaddedfields
        
        For i = 1 To .ListItems.Count
          Set X = .ListItems.Item(i)
          If .ListItems.Item(i).Selected = True Then
            store = X.SubItems(1)
            For j = 1 To Me.lstfields.ListItems.Count
              Set X = Me.lstfields.ListItems.Item(j)
              If Me.lstfields.ListItems.Item(j) = store Then
                X.SubItems(1) = "-"
              End If
            Next
           .ListItems.Remove (i)
            Exit For
          End If
        Next
        
      End With
    End If

Exit Sub
jump:
MsgBox Err.Description, vbCritical
End Sub

Private Sub optasp_Click()
  heading.Caption = "Specify The Path Of Your Virtual Directory"
End Sub

Private Sub optdeselect_Click()

  If Me.lstfields.ListItems.Count >= 1 Then
    For i = 1 To Me.lstfields.ListItems.Count
      Me.lstfields.ListItems.Item(i).Checked = False
    Next
  End If
  
End Sub

Private Sub opthtml_Click()
   heading.Caption = "Specify The Path To Store The Files"
End Sub

Private Sub optselect_Click()

  If Me.lstfields.ListItems.Count >= 1 Then
    For i = 1 To Me.lstfields.ListItems.Count
      Me.lstfields.ListItems.Item(i).Checked = True
    Next
  End If
  
End Sub


