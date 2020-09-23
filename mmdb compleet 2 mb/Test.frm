VERSION 5.00
Begin VB.Form Formreg3 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Geregiseerde ......."
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Start MMDB"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   -240
      Picture         =   "Test.frx":0000
      ScaleHeight     =   1065
      ScaleWidth      =   4545
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unregister"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Geregistreerde: Unregistered"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "Formreg3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Kill ("C:\Windows\System\wind32.dat")
MsgBox ("All Done!"), vbInformation, ("Done!")
SaveSetting "Win32", "Startup", "counter", 1
       SaveSetting "Win32", "Startup", "Started", Format(Date, "mm dd yyyy")
       SaveSetting "Win32", "Startup", "Last Used", Format(Date, "mm dd yyyy")
End
End Sub

Private Sub Command2_Click()
MsgBox ("You are in the registered Version!"), vbInformation, ("Registered Version!")
End Sub

Private Sub Command3_Click()
Unload Me
Load FrmMain
FrmMain.Show
End Sub

Private Sub Form_Load()

'######## This checks if the program has run
'and if the user deleted the registry to start his time over again!

If Label3.Caption = ("0") Then


Open "C:\windows\system\windat2.dat" For Input As #1
Input #1, started
Close #1

' If the program has run before and the counter has 0 when it should have more, then
If started = "yes" And GetSetting("Win32", "Startup", "counter", "") = "" Then
    MsgBox "NO! Deleting the Registry Entries Won't do Crap you stupid Hacker!", vbInformation + vbOKOnly, "Ooops"
     End
     Exit Sub
    End If

End If



'#### End Of Registry code














On Error GoTo 55
Open "C:\windows\system\wind32.dat" For Input As #1
Input #1, final1
Input #1, fname
Close #1
Dim Code1 As Single

If Len(fname) < 4 Then
    MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
    Exit Sub
End If

Label1.Caption = ("Registered To: ") & fname

If Label1.Caption = "Registered To:" Then
Label1.Caption = ("Registered To: Unregistered")
End If





For i = 1 To Len(fname) - 1
    Code1 = Format(Asc(Right(fname, Len(fname) - i)) * 2 + (31 / i) + (i + 3 / 7), "#.#")
    zip = zip & Code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 7), "#00")
    final = final & Code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(fname)
If final1 = final And Len(fname) >= 4 Then
   'ENABLE FUNCTIONS IF REGISTERED
    Command2.Enabled = True
    Command1.Enabled = True
    Label3.Caption = ("1")
Else
55
    Me.Hide
    Formreg1.Show
    Exit Sub
End If





End Sub

