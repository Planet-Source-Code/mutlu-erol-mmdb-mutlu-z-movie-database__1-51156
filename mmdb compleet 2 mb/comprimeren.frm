VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form comprimeren 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Database Comprimeren "
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "comprimeren.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   112.462
   ScaleMode       =   0  'User
   ScaleWidth      =   122.167
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   26
      Top             =   240
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
         TabIndex        =   27
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15480
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   375
         Left            =   10200
         TabIndex        =   51
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   375
         Left            =   10080
         TabIndex        =   50
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000009&
         Height          =   285
         Left            =   6240
         TabIndex        =   49
         Text            =   "1                   5                       10"
         Top             =   5640
         Width           =   2295
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "comprimeren.frx":09CA
         Height          =   7470
         Left            =   12360
         TabIndex        =   46
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   13176
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483642
         ForeColor       =   65535
         ListField       =   "filmnaam"
         Object.DataMember      =   "connect"
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000A&
         Caption         =   "&Bewerken"
         Height          =   255
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   8760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "&Nieuwe film"
         Height          =   315
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   8400
         Width           =   975
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1200
         Picture         =   "comprimeren.frx":09E9
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   42
         Top             =   6000
         Width           =   375
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2280
         MouseIcon       =   "comprimeren.frx":0D54
         Picture         =   "comprimeren.frx":105E
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   41
         Top             =   6000
         Width           =   375
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   13680
         Picture         =   "comprimeren.frx":13CB
         ScaleHeight     =   825
         ScaleWidth      =   1665
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Afsluiten"
         Height          =   195
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   600
         Width           =   135
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9360
         Top             =   6840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3975
         Left            =   4080
         ScaleHeight     =   3915
         ScaleWidth      =   2955
         TabIndex        =   30
         Top             =   4320
         Width           =   3015
         Begin VB.PictureBox picpic 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2040
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   31
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   8520
         Top             =   7200
      End
      Begin VB.TextBox txtfilmnaam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Vul hier de naam van de film"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox txtacteurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Vul hier de naam van de acteurs"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtgenre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Actie ,horror,thriller.etx"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox txttype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "DivX of DVD of VHS"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtspeelduur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Vul hier de speelduur van de film"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtleefdtijdsgrens 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   " **********"
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox txtuitgeleend 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "Aan wie heb je hem uitgeleend?"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox txtmisc 
         BackColor       =   &H80000006&
         ForeColor       =   &H00C0FFC0&
         Height          =   1875
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "comprimeren.frx":1DBE
         Top             =   7200
         Width           =   5895
      End
      Begin VB.TextBox txtjaar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Welke jaar?"
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000012&
         Caption         =   "<<"
         Height          =   105
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   5520
         Width           =   135
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000012&
         Caption         =   ">>"
         Height          =   105
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   5520
         Width           =   135
      End
      Begin VB.TextBox txtpost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1695
         TabIndex        =   13
         Text            =   "1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3960
         Picture         =   "comprimeren.frx":1DD8
         ScaleHeight     =   1065
         ScaleWidth      =   5025
         TabIndex        =   33
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Acteurs                    :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Genre                      :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Type film                  :            "
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   22
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "Speelduur                :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Rating  (max/10)     :"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         Caption         =   "Jaar                         :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Uitgeleend aan        :"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000007&
         Caption         =   "Films die ik heb"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   13320
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Film cover"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   32
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Film naam"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Film info"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   6840
         Width           =   1695
      End
      Begin VB.Label lblcreated 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   4200
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   9120
      Width           =   15255
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3840
         Picture         =   "comprimeren.frx":3F1D
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   48
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2040
         Picture         =   "comprimeren.frx":4A3A
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   47
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   11880
         Picture         =   "comprimeren.frx":54C6
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   38
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         Picture         =   "comprimeren.frx":5DFB
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   37
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5520
         Picture         =   "comprimeren.frx":6811
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   7320
         Picture         =   "comprimeren.frx":71F9
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   35
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9480
         Picture         =   "comprimeren.frx":7BD2
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   34
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Verwijder alles"
         Height          =   195
         Left            =   12360
         TabIndex        =   28
         Top             =   360
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000A&
         Caption         =   "&Verwijder"
         Height          =   255
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Zoeken"
         Height          =   195
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   4005
      _ExtentX        =   7064
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
         EndProperty
      EndProperty
   End
   Begin VB.Menu Afsluiten 
      Caption         =   "Comprimeren en daarna afsluiten"
   End
End
Attribute VB_Name = "comprimeren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'Kill App.Path & "\compact.mdb"
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\compact.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\compact.mdb", App.Path & "\mutlu.mdb")
Kill App.Path & "\compact.mdb"
End If
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""

End Sub

Private Sub Bewerken_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "     van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
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

Private Sub Command1_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "         van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
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

Private Sub Command10_Click()
Kill App.Path & "\compact.mdb"
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtleefdtijdsgrens.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
Command8.Enabled = False
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End If
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

Private Sub Command4_Click()
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

Private Sub Command5_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "     van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
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

Private Sub Command6_Click()
If FrmFind.Visible = False Then
FrmFind.Show 0, Me
Else
Unload FrmFind
End If
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
'Kill App.Path & "\compact.mdb"
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\compact.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\compact.mdb", App.Path & "\mutlu.mdb")
Kill App.Path & "\compact.mdb"
End If
Unload FrmFind
Unload Me
End
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
End Sub






Private Sub Command9_Click()
Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\compact.mdb")
Kill App.Path & "\mutlu.mdb"
Call FileCopy(App.Path & "\compact.mdb", App.Path & "\mutlu.mdb")

End Sub

Private Sub cover_Click()
If editing = True Then
    
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.FileName)
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
If Form1.Visible = False Then
Form1.Show 0, Me
Else
Unload Form1
End If

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

postnr = 1
If Dir(App.Path & "\mutlu.mdb") <> "" Then GoTo continue
CreateNewDB "mutlu.mdb"
continue:

Set Db = OpenDatabase(App.Path & "\mutlu.mdb")

Set Rs = Db.OpenRecordset("filmlijst")

If Rs.RecordCount > 0 Then
    'Rs.MoveFirst
    'txtfilmnaam = Trim(Rs.Fields(0))
    'txtacteurs = Trim(Rs.Fields(1))
    'txtgenre = Trim(Rs.Fields(2))
    'txttype = Trim(Rs.Fields(3))
    'txtspeelduur = Trim(Rs.Fields(4))
    'txtjaar = Trim(Rs.Fields(5))
    'txtleefdtijdsgrens = Trim(Rs.Fields(6))
    'txtuitgeleend = Trim(Rs.Fields(7))
    'txtmisc = Trim(Rs.Fields(8))
    'lblcreated.Caption = "Post skapad " & Trim(Rs.Fields(9))
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
    visapost (1)
    'StatusBar1.Panels(1).Text = "Det finns" & Str(Rs.RecordCount) & " poster i databasen. Visar post 1."
    'txtpost = "1"
    
    Open App.Path & "\mutlu.mdb" For Binary As #1
    g = LOF(1)
    Close #1
'StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
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
    'Open the RecordSet to get the Categories List
    Set Rs = Db.OpenRecordset("filmlijst")
    On Error Resume Next
        Rs.AddNew
        
        Rs.Fields(0) = Trim(txtfilmnaam)
        Rs.Fields(1) = Trim(txtacteurs)
        Rs.Fields(2) = Trim(txtgenre)
        Rs.Fields(3) = Trim(txttype)
        Rs.Fields(4) = Trim(txtspeelduur)
        Rs.Fields(5) = Trim(txtjaar)
        Rs.Fields(6) = Trim(txtleefdtijdsgrens)
        Rs.Fields(7) = Trim(txtuitgeleend)
        Rs.Fields(8) = Trim(txtmisc)
        Rs.Fields(9) = Trim(Now)
        
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
        txtleefdtijdsgrens.Locked = True
        txtuitgeleend.Locked = True
        txtmisc.Locked = True
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
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "         van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
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




Private Sub Picture10_Click()
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
'Kill App.Path & "\compact.mdb"
'Call CompactDatabase(App.Path & "\mutlu.mdb", App.Path & "\compact.mdb")
'Kill App.Path & "\mutlu.mdb"
'Call FileCopy(App.Path & "\compact.mdb", App.Path & "\mutlu.mdb")
'Kill App.Path & "\compact.mdb"
End If
Unload FrmFind
Unload Me
End

End Sub

Private Sub Picture11_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If

End Sub

Private Sub Picture12_Click()
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

Private Sub Picture13_Click()
If editing = True Then
    
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.FileName)
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

Private Sub Picture2_DblClick()
If editing = True Then
    
    CommonDialog1.ShowOpen
    picpic.Visible = True
    picpic.Picture = LoadPicture(CommonDialog1.FileName)
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

Private Sub Picture4_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    picpic.Visible = False
    picpic.Picture = LoadPicture
    Picture2.Cls
    Picture2.Print
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "         van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
    txtfilmnaam = ""
    txtacteurs = ""
    txtgenre = ""
    txttype = ""
    txtspeelduur = ""
    txtjaar = ""
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
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

Private Sub Picture5_Click()
If editing = False Then
    Command3.Enabled = False
    Command4.Enabled = False
    txtpost.Locked = True
    If picpic.Picture = 0 Then
    Picture2.Cls
    Picture2.Print
    Picture2.Print "      Klik hier om een cover toe te voegen"
    Picture2.Print "       of druk op ctrl+v om te plakken"
    Picture2.Print "     van een plaatje via paint"
    Picture2.Print
    Picture2.Print "        De grootte van de plaatje"
    Picture2.Print "    moet niet groter zijn dan 193x258"
    Picture2.Print "        pixels <M.E>.."
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
    txtleefdtijdsgrens.Locked = False
    txtuitgeleend.Locked = False
    txtmisc.Locked = False
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

Private Sub Picture6_Click()
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtleefdtijdsgrens.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
StatusBar1.Panels(1).Text = "Er zijn geen film in de database."
Command8.Enabled = False
StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End If

End Sub

Private Sub Picture7_Click()
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""

End Sub

Private Sub Picture8_Click()
If Form1.Visible = False Then
Form1.Show 0, Me
Else
Unload Form1
End If

End Sub

Private Sub Picture9_Click()
If FrmFind.Visible = False Then
FrmFind.Show 0, Me
Else
Unload FrmFind
End If

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
        Rs.Fields(6) = Trim(txtleefdtijdsgrens)
        Rs.Fields(7) = Trim(txtuitgeleend)
        Rs.Fields(8) = Trim(txtmisc)
        Rs.Fields(9) = Trim(Now)
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
        txtleefdtijdsgrens.Locked = True
        txtuitgeleend.Locked = True
        txtmisc.Locked = True
        Open App.Path & "\mutlu.mdb" For Binary As #1
        g = LOF(1)
        Close #1
        StatusBar1.Panels(2).Text = "Grootte van de database : " & Format(g, "###,###,###,##0") & " k"

        visapost post
        
End Sub



Private Sub uitgeleendaan_Click()
If Form1.Visible = False Then
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
    txtleefdtijdsgrens = ""
    txtuitgeleend = ""
    txtmisc = ""
    txtfilmnaam.Locked = True
    txtacteurs.Locked = True
    txtgenre.Locked = True
    txttype.Locked = True
    txtspeelduur.Locked = True
    txtjaar.Locked = True
    txtleefdtijdsgrens.Locked = True
    txtuitgeleend.Locked = True
    txtmisc.Locked = True
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
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = LoadPicture(App.Path & "\but_add1.jpg")

End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.Picture = LoadPicture(App.Path & "\but_add.jpg")

End Sub


Private Sub zoekfilmapart_Click()
If frmList.Visible = False Then
frmList.Show 0, Me
Else
Unload frmList
End If

End Sub
