VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M.E.'z Start Kazaa"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Form3.frx":09CA
   ScaleHeight     =   3750
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Klik hier om Kazaa te starten"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H80000006&
      Class           =   "Package"
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form3.frx":C898
      SourceDoc       =   "C:\Program Files\KaZaA\Kazaa.exe"
      TabIndex        =   0
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
