VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "------ M.E.'z DivX BIOS ------ "
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   15270
   FillColor       =   &H00FFFFC0&
   FillStyle       =   0  'Solid
   Icon            =   "ddiary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ddiary.frx":09CA
   ScaleHeight     =   10725
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Welkom bij -[M.E'z DivX Bios]- Neem uw plaats in, en selecteer uw film door op het scherm te klikken"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   9960
      Width           =   7455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9360
      Top             =   4320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Klik hier om een film te selecteren die je wilt kijken"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ddiary.frx":C5F7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   45
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   10080
      Picture         =   "ddiary.frx":11C05
      ScaleHeight     =   105
      ScaleWidth      =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   15
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H80000001&
      Height          =   5415
      Left            =   2760
      ScaleHeight     =   5415
      ScaleWidth      =   9975
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "M.E.'z DivX Bios"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1800
         TabIndex        =   6
         Top             =   2160
         Width           =   6375
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   6015
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   9735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   0
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      Height          =   4215
      Left            =   3600
      Top             =   1080
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


' This flag is set when the user chooses Cancel.
Dim CancelFlag
Dim i As Double
Dim a(100) As String
Dim add As Double
Dim d As Integer
Dim e As Double
Dim su As Double
Dim s As Integer
Dim g As Double
Dim mul As Double
Dim f As Integer
Dim m As Integer
Dim n As Double
Dim div As Double
Dim o As Integer
Dim p As Double
Dim sq As Double

Dim si As Double
Dim co As Double
Dim ta As Double

Private Sub Command1_Click()
Memo = 1
Form3.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command26_Click()
text1.Text = "1"
a(i) = (text1.Text)

i = i + 1
For b = 0 To i - 1
    c = c + a(b)
    text1.Text = c
Next b

End Sub

Private Sub Command10_Click()
text1.Text = "0"
a(i) = (text1.Text)

i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command11_Click()
If f = 0 Then
text1.Text = "."
a(i) = (text1.Text)

i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
f = 1
Else
MsgBox "DON't TRY TO USE TWO DECIMAL POINTS", vbCritical, "ERROR"
End If
End Sub

Private Sub Command12_Click()
f = 0
add = text1.Text
text1.Text = "0."

d = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command13_Click()
f = 0
su = text1.Text
text1.Text = "0."

s = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command14_Click()
f = 0
mul = text1.Text
text1.Text = "0."

m = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command15_Click()
f = 0
div = text1.Text
text1.Text = "0."

o = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command16_Click()
For b = 0 To i - 1
a(b) = 0
Next b
i = 0
text1.Text = "0."
End Sub

Private Sub Command17_Click()
If d = 1 Then
e = text1.Text + add
text1.Text = e
d = 0
ElseIf s = 1 Then
g = su - text1.Text
text1.Text = g
s = 0
ElseIf m = 1 Then
n = mul * text1.Text
text1.Text = n
m = 0
ElseIf o = 1 Then
If (text1.Text > 0) Then
p = div / text1.Text
text1.Text = p
o = 0
ElseIf (text1.Text) = 0 Then
MsgBox ("INVALID CAN'T DIVIDE BY 0"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
text1.Text = "0."
End If
End If

For b = 0 To i - 1
a(b) = 0
Next b
End Sub

Private Sub Command18_Click()
If (text1.Text) > 0 Then
sq = Sqr(text1.Text)
text1.Text = sq

ElseIf text1.Text = 0 Then
MsgBox ("INVALID: NOT DEFINED"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
text1.Text = "0."
End If
End Sub

Private Sub Command19_Click()
si = Sin(text1.Text)
text1.Text = si
End Sub

Private Sub Command25_Click()
text1.Text = "2"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command20_Click()
co = Cos(text1.Text)
text1.Text = co
End Sub

Private Sub Command21_Click()
ta = Tan(text1.Text)
text1.Text = ta
End Sub

Private Sub Command24_Click()
text1.Text = "3"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command23_Click()
text1.Text = "4"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command22_Click()
text1.Text = "5"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command27_Click()
Dim mytime
mytime = InputBox("ENTER NEW TIME")
Time = mytime
End Sub
Private Sub Command28_Click()
Form8.Show
End Sub

Private Sub Command29_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command3_Click()
For z = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height - z
Next z
Picture1.Visible = False
For V = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height + V
Next V
On Error GoTo form1
CommonDialog1.DialogTitle = "Load Media"
CommonDialog1.CancelError = True
CommonDialog1.Filter = "AVI Files|*.avi|WAV Files|*.wav|MIDI files|*.mid|WMV Files|*.wmv|All Files|*.*"
CommonDialog1.ShowOpen
MediaPlayer1.Open (CommonDialog1.filename)
mnuPlay.Enabled = True
mnuStop.Enabled = True
mnuPause.Enabled = True
mnuRewind.Enabled = True
form1:
form1.Show
End Sub


Private Sub Command30_Click()
Form10.Show
End Sub

Private Sub Command31_Click()
Form9.Show
End Sub



Private Sub Dial(Number$)
    Dim DialString$, FromModem$, dummy

    ' AT is the Hayes compatible ATTENTION command and is required to send commands to the modem.
    ' DT means "Dial Tone." The Dial command uses touch tones, as opposed to pulse (DP = Dial Pulse).
    ' Numbers$ is the phone number being dialed.
    ' A semicolon tells the modem to return to command mode after dialing (important).
    ' A carriage return, vbCr, is required when sending commands to the modem.
    DialString$ = "ATDT" + Number$ + ";" + vbCr

    ' Communications port settings.
    ' Assuming that a mouse is attached to COM1, CommPort is set to 2
    MSComm1.CommPort = 2
    MSComm1.Settings = "9600,N,8,1"
    
    ' Open the communications port.
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    ' Flush the input buffer.
    MSComm1.InBufferCount = 0
    
    ' Dial the number.
    MSComm1.Output = DialString$
    
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, "OK") Then
             ' Notify the user to pick up the phone.
             Beep
             MsgBox "Please pick up the phone and either press Enter or click OK", , "Digital Diary"
             Exit Do
          End If
       End If
        
       ' Did the user choose Cancel?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Disconnect the modem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Close the port.
    MSComm1.PortOpen = False
End Sub
Private Sub Command33_Click()
    ' CancelFlag tells the Dial procedure to exit.
   CancelFlag = True
   Command33.Enabled = False
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Label6.Visible = False
CommonDialog2.Filter = "All Files|*.*|BMP Files|*.bmp|JPEG files|*.jpeg|GIF Files|*.gif|JPG Files|*.jpg|ICO files|*.ico|CUR files|*.cur"
CommonDialog2.ShowOpen
Image1.Picture = LoadPicture(CommonDialog2.filename)
End Sub

Private Sub Command6_Click()
text1.Text = "6"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command7_Click()
text1.Text = "7"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command8_Click()
text1.Text = "8"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub

Private Sub Command9_Click()
text1.Text = "9"
a(i) = (text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 text1.Text = c
 Next b
End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload form1
End Sub

Private Sub Label1_Click()
 MciSendString "set CDAudio door open", vbNullString, 0&, 0&

For z = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height - z
Next z
Picture1.Visible = False
For V = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height + V
Next V
On Error GoTo form1
CommonDialog1.DialogTitle = "Selecteer een film"
CommonDialog1.CancelError = True
CommonDialog1.Filter = "AVI Files|*.avi|WAV Files|*.wav|MPG files|*.mpg|WMV Files|*.wmv|MPEG files|*.mpeg|All Files|*.*"
CommonDialog1.ShowOpen
MediaPlayer1.Open (CommonDialog1.filename)
mnuPlay.Enabled = True
mnuStop.Enabled = True
mnuPause.Enabled = True
mnuRewind.Enabled = True
Unload form1
form1:
form1.Show
Unload form1
End Sub


Private Sub info_Click()
If frmAbout1.Visible = False Then
frmAbout1.Show 0, Me
Else
Unload frmAbout1
End If

End Sub

Private Sub Picture1_Click()
 MciSendString "set CDAudio door open", vbNullString, 0&, 0&

For z = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height - z
Next z
Picture1.Visible = False
For V = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height + V
Next V
On Error GoTo form1
CommonDialog1.DialogTitle = "Selecteer een film"
CommonDialog1.CancelError = True
CommonDialog1.Filter = "AVI Files|*.avi|WAV Files|*.wav|MPG files|*.mpg|WMV Files|*.wmv|MPEG files|*.mpeg|All Files|*.*"
CommonDialog1.ShowOpen
MediaPlayer1.Open (CommonDialog1.filename)
mnuPlay.Enabled = True
mnuStop.Enabled = True
mnuPause.Enabled = True
mnuRewind.Enabled = True
Unload form1
form1:
form1.Show
Unload form1
End Sub

Private Sub Picture2_Click()
If FrmMain.Visible = False Then
FrmMain.Show 0, Me
Else
Unload form1
End If

End Sub

Private Sub CDClose_Click()
MciSendString "set CDAudio door closed", vbNullString, 0&, 0&
End Sub
Private Sub CDOpen_Click()
 MciSendString "set CDAudio door open", vbNullString, 0&, 0&
End Sub
