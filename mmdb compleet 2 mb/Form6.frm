VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   LinkTopic       =   "Form6"
   ScaleHeight     =   7695
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Terug"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   7320
      Width           =   735
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "Form6.frx":0000
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12488
      _Version        =   393216
      Appearance      =   0
      BackColor       =   0
      ForeColor       =   255
      ListField       =   "filmnaam"
      Object.DataMember      =   "connect"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

