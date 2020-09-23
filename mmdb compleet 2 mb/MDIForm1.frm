VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000012&
   Caption         =   "M.E.'z Internet browser"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12060
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":09CA
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu home 
      Caption         =   "Enter   --> Mutlu'z MovieDataBase <--  Online"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub home_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://members.home.nl/m.erol/mmdb.htm"
    frmB.Show

End Sub
