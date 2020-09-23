VERSION 5.00
Begin VB.Form Back 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   ScaleHeight     =   8655
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "HTML converter"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   11400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   3
      Height          =   8895
      Left            =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   11520
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Convert.Show

End Sub

'Sub Colors()
'Dim i
'With Me
 '.DrawStyle = vbInsideSolid
 '.DrawMode = vbCopyPen
 '.ScaleMode = vbPixels
 '.DrawWidth = 2
 '.ScaleHeight = 240
 'For i = 0 To 255
  'Line (0, i)-(Screen.Width, i - 1), RGB(0, 0, 255 - i), B
 'Next
'End With
'End Sub

 Private Sub Form_Paint()
'Colors
'Convert.Show
 End Sub
