VERSION 5.00
Begin VB.Form tabellen 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Tabellen"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   Icon            =   "tabellen.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   8415
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox serienummer 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "serienummer"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox ondertiteling 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "ondertiteling"
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "tabellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
