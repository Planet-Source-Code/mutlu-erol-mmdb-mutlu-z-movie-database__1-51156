VERSION 5.00
Begin VB.Form Formreg2 
   BackColor       =   &H80000008&
   Caption         =   "Registreren"
   ClientHeight    =   2670
   ClientLeft      =   3375
   ClientTop       =   2490
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Registreren"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      MaxLength       =   17
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Serienmr #:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bedrijf:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Naam:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Formreg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox ("Vul alles in!"), vbInformation, ("Registration")
Exit Sub
End If


If Len(Text1.Text) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
    Exit Sub
End If

If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
    MsgBox "Registratie Klopt niet! Probeer het opnieuw", vbCritical, ("Registration")
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    Code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (31 / i) + (i + 3 / 7), "#.#")
    zip = zip & Code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 7), "#00")
    final = final & Code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
If Text2.Text = final Then
    Close #1
    Open "C:\windows\system\wind32.dat" For Output As #1
    Write #1, final
    Write #1, Text1.Text
    Close #1
    Command1.Caption = "Geregistreerd!"
    MsgBox "Bedankt voor het registreren van mijn programma! Start de programma opnieuw op.", vbInformation + vbOKOnly, "Geregistreerd"
    End
Else
    MsgBox "Registration Failed. Please check your information", vbCritical, ("Registration")
End If

End Sub

Private Sub Command2_Click()
Formreg2.Hide
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Formreg1.Show
End Sub

Private Sub Form_Load()
DoG Formreg2
End Sub

