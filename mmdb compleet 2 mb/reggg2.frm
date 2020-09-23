VERSION 5.00
Begin VB.Form Formreg1 
   BackColor       =   &H80000007&
   Caption         =   "14 dagen TRIAL periode"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   Icon            =   "reggg2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   360
      Picture         =   "reggg2.frx":27A2
      ScaleHeight     =   1065
      ScaleWidth      =   4785
      TabIndex        =   6
      Top             =   0
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "X"
      TabIndex        =   5
      Text            =   "yes"
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   2640
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Doorgaan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   3000
      Picture         =   "reggg2.frx":48E7
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registeer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1320
      Picture         =   "reggg2.frx":4F75
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label Label3 
      Height          =   135
      Left            =   5640
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You are on day 0 of your 14  day trial period."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label lblcnt 
      Height          =   15
      Left            =   4920
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "Formreg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'VerifyKeyCode Text1.Text, Text2.Text
'' End Sub

Private Sub Form_Load()

DoG form1



'This is the code for the trial version

Dim startdate As String
Dim differenceofdate
Dim TRACEDATE As String
Dim newdate
Dim chk
       Label3.Caption = GetSetting("Win32", "startup", "counter", "")
    If Label3.Caption = "" Then
 
   
   
        SaveSetting "Win32", "Startup", "counter", 1
       SaveSetting "Win32", "Startup", "Started", Format(Date, "mm dd yyyy")
       SaveSetting "Win32", "Startup", "Last Used", Format(Date, "mm dd yyyy")
       
       Label3.Caption = "1"
       Label2.Caption = "You are on day 1 of your 14  day trial period."
           '######   This Makes sure that that the user won't delete the registry!
          
    Close #1
    Open "C:\windows\system\windat2.dat" For Output As #1
    Write #1, Text1.Text
    Close #1
    '######  End of checking registry
    End If

    If Label3.Caption >= 14 Then
       MsgBox "Trial Period Has Expired!" & Chr(10) + Chr(13) & "Please Register or unistall this program.", vbCritical, "Trial Period"
      Label2.Caption = ("Trial Period Expired.")
      Timer1.Enabled = True
      Shape1.FillColor = vbRed
Label2.ForeColor = vbBlack
       
 Else
       TRACEDATE = GetSetting("Win32", "Startup", "Last Used", "")
chk = DateDiff("d", CDate(TRACEDATE), Now)
       If chk < 0 Then 'CHECK IF THE DATE WAS CHANGE which is lesser than the PREVIOUS DATE WHERE THE SYSTEM USED.
          MsgBox "Your system date is invalid." & Chr(10) + Chr(13) & "Please change it right now or else you will nolonger use this software anymore!!", vbCritical, "Invalid date"
                End
       Else
      startdate = GetSetting("Win32", "Startup", "Started", "")
       differenceofdate = DateDiff("d", startdate, Now)
         If differenceofdate <> 0 Then
                    Label3.Caption = differenceofdate + 1
        SaveSetting "Win32", "Startup", "Last Used", Format(Now, "MM DD YYYY")
               SaveSetting "Win32", "Startup", "counter", differenceofdate + 1
                End If
                If differenceofdate = 0 Then
                Label3.Caption = GetSetting("Win32", "Startup", "Counter", "")
               If Label3.Caption >= 14 Then
       MsgBox "Trial Period Has Expired!" & Chr(10) + Chr(13) & "Please Register or unistall this program.", vbCritical, "Trial Period"
      Label2.Caption = ("Trial Period Expired.")
      Timer1.Enabled = True
      Shape1.FillColor = vbRed
Label2.ForeColor = vbBlack
                
                End If
       End If
  End If
  
  '############Beginning Of Start Code
'  Label1.Caption = "Thank You For Using " & App.Title

Label2.Caption = "You are on day " & Label3.Caption & " of your 14  day trial period."

If Label3.Caption >= "7" Then
Shape1.FillColor = vbYellow
Label2.ForeColor = vbBlack
'############End Of Start Code
End If
  '################  End Of Trial Code
  End If
  
End Sub

Private Sub Label4_Click()
Formreg2.Show
Formreg1.Hide
End Sub

Private Sub Label5_Click()
If GetSetting("Win32", "Startup", "counter", "") >= "7" Then
MsgBox ("How dare you try to hack this program. I don't Think So!!!"), vbCritical, ("Hack Proof Protection")
End
Exit Sub

Else
FrmMain.Show
Me.Visible = False
End If


End Sub

Private Sub Timer1_Timer()
Label5.Enabled = False
End Sub
