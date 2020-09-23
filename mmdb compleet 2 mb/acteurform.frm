VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Acteurformulier 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acteuren tool Created by Mutlu Erol 21 november"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6990
   Icon            =   "acteurform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "ok"
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MMDB acteurs database"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Converteer naar hoofd menu"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   5040
      TabIndex        =   4
      Top             =   7800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Voeg acteur"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Text            =   "Gekozen acteurs"
      Top             =   1320
      Width           =   2535
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "acteurform.frx":09CA
      Height          =   4545
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8017
      _Version        =   393216
      ListField       =   "First"
      BoundColumn     =   "First"
      Object.DataMember      =   "Command1"
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Sluit MMDB eerst af voordat je Acteuren database opstart!!!!!!!"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Converteren liep succesvol"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Voeg hier zelf een acteur naam in, als de acteur niet in je lijst staat"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Klanten tool by M.E."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Meer dan  acteur ??? Copy & Paste de gekozen acteurs naar dit vakje"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Toegevoegde acteur"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Beschikbare acteurs in de database"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Acteurformulier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1 = DataList1
Combo1.AddItem Text1
List1.AddItem Text1


End Sub

Private Sub Command2_Click()
FrmMain.txtacteurs = Text2
Label6.Visible = True

End Sub

Private Sub info_Click()

End Sub

Private Sub Command4_Click()
Unload Me
End Sub
