VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMainnew 
   BackColor       =   &H80000006&
   Caption         =   "------[MMDB Light 2003 ©]-------"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11895
   Icon            =   "frmMainnew.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4290
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "22-1-2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "3:11"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":09CA
            Key             =   "circle_f1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":0CE4
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":0DF6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":0F08
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":101A
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":112C
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":123E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":1350
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":1462
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":1574
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":1686
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":1798
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainnew.frx":18AA
            Key             =   "Sort Descending"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "1037"
            ImageKey        =   "circle_f1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "1038"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "1040"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Vorige film"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Volgende film"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Zoeken"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort Ascending"
            Object.ToolTipText     =   "1048"
            ImageKey        =   "Sort Ascending"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort Descending"
            Object.ToolTipText     =   "1049"
            ImageKey        =   "Sort Descending"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Start"
      WindowList      =   -1  'True
   End
   Begin VB.Menu bestand 
      Caption         =   "MMDB"
      Begin VB.Menu afsluiten 
         Caption         =   "Afsluiten"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu bewerken 
      Caption         =   "Bewerken"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu vorigefilm 
         Caption         =   "Vorige film"
         Shortcut        =   ^N
      End
      Begin VB.Menu volgendefilm 
         Caption         =   "Volgende film"
         Shortcut        =   ^M
      End
      Begin VB.Menu sorteermenu 
         Caption         =   "Sorteer Menu"
      End
      Begin VB.Menu zoeken 
         Caption         =   "Zoeken"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu bios 
      Caption         =   "Bios"
      Begin VB.Menu filmbekijken 
         Caption         =   "Film bekijken"
         Shortcut        =   ^T
      End
      Begin VB.Menu cdromopen 
         Caption         =   "Open CD-rom"
      End
      Begin VB.Menu cdromclose 
         Caption         =   "Sluit CD-rom"
      End
   End
   Begin VB.Menu boekhouding 
      Caption         =   "Filmlijst"
      Begin VB.Menu filmlijsthtml 
         Caption         =   "Filmlijst HTML"
      End
      Begin VB.Menu filmlijsthtmlconverter 
         Caption         =   "Filmlijst HTML Converter"
      End
      Begin VB.Menu geavanceerd 
         Caption         =   "Geavanceerde film info"
      End
   End
   Begin VB.Menu covers 
      Caption         =   "Covers"
      Begin VB.Menu coverdatabase 
         Caption         =   "MMDB Cover Database"
      End
      Begin VB.Menu covereditor 
         Caption         =   "Cover editor"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu skin 
      Caption         =   "Skins"
      Begin VB.Menu skin2 
         Caption         =   "Original skin"
      End
      Begin VB.Menu skin1 
         Caption         =   "Black skin"
      End
      Begin VB.Menu banderas 
         Caption         =   "Antonio banderas skin"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu helpindex 
         Caption         =   "Help-onderwerpen"
         Enabled         =   0   'False
      End
      Begin VB.Menu info 
         Caption         =   "Info"
      End
      Begin VB.Menu about 
         Caption         =   "About MMDB"
      End
   End
   Begin VB.Menu mnuViewWebBrowser 
      Caption         =   "Internet"
      Begin VB.Menu zoektrailer 
         Caption         =   "Zoek Trailer"
         Visible         =   0   'False
      End
      Begin VB.Menu biosnl 
         Caption         =   "www.bios.nl"
      End
      Begin VB.Menu videoland 
         Caption         =   "www.videoland.nl"
      End
      Begin VB.Menu imdb 
         Caption         =   "www.imdb.com"
      End
      Begin VB.Menu amazon 
         Caption         =   "www.amazon.com"
      End
   End
   Begin VB.Menu updatee 
      Caption         =   "Update"
   End
   Begin VB.Menu medesign 
      Caption         =   "M.E.Design website"
   End
   Begin VB.Menu mmdbregistreren 
      Caption         =   "MMDB registreren"
   End
End
Attribute VB_Name = "frmMainnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub about_Click()
frmAbout99.Show

End Sub

Private Sub Afsluiten_Click()
Unload Me

End Sub

Private Sub amazon_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.amazon.com"
    frmB.Show
End Sub

Private Sub banderas_Click()
FrmMain.Image12.Visible = False
FrmMain.Image13.Visible = True
FrmMain.Image13.Picture = LoadPicture(App.Path & "\banderas.jpg")

End Sub

Private Sub bewerkenopslaan_Click()
End Sub

Private Sub biosnl_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.bios.nl"
    frmB.Show
End Sub

Private Sub cdromclose_Click()
MciSendString "set CDAudio door closed", vbNullString, 0&, 0&
End Sub

Private Sub cdromopen_Click()
 MciSendString "set CDAudio door open", vbNullString, 0&, 0&
End Sub

Private Sub download_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "C:\Program Files\Mutlu'z Movie database\filmlijst.html"
    frmB.Show
End Sub

Private Sub coverdatabase_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://members.lycos.nl/tatlicocuk22"
    frmB.Show
End Sub

Private Sub coverprint_Click()

End Sub

Private Sub filmbekijken_Click()
Dim f As New form1
            f.Show
End Sub

Private Sub filmlijsthtml_Click()
' MsgBox "NVT in deze versie"

'shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\first.bat", vbNullString, vbNullString, SW_Shownormal)

Dim frmB As New frmBrowser
 frmB.StartingAddress = "C:\Program Files\Mutlu'z Movie database\HTML1\filmlijst.html"
   frmB.Show
End Sub

Private Sub filmlijsthtmlconverter_Click()
' MsgBox "NVT in deze versie!"
' shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\check.bat", vbNullString, vbNullString, SW_Shownormal)
' MsgBox "Vorige html-bestanden zijn gewist"
'shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\kill.bat", vbNullString, vbNullString, SW_Shownormal)
'MsgBox "Nieuwe Filmlijst html-bestand is aangemaakt "
'shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\renamehtml.bat", vbNullString, vbNullString, SW_Shownormal)
'MsgBox "Html converter word gestart"

Dim f As New Back
Convert.Show
'           f.Show

End Sub

Private Sub filmszoeken_Click()
'Dim f As New kazaa
 '           f.Show
End Sub

Private Sub geavanceerd_Click()
Dim f As New filminfo
            f.Show
End Sub

Private Sub helpindex_Click()
MsgBox "N.V.T. in deze versie!"
End Sub

Private Sub home_Click()
End Sub

Private Sub imdb_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.imdb.com"
    frmB.Show
End Sub

Private Sub info_Click()
Dim f As New Form2
            f.Show
End Sub

Private Sub info2_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://members.lycos.nl/tatlicocuk21"
    frmB.Show

End Sub

Private Sub MDIForm_Load()
    LoadResStrings Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
End Sub





Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuHelpAbout_Click()
Dim f As New form1
    f.Show
End Sub

Private Sub medesign_Click()
Dim frmB As New frmBrowser
frmB.StartingAddress = "http://members.lycos.nl/tatlicocuk21"
frmB.Show
End Sub

Private Sub mmdbregistreren_Click()
Dim frmB As New frmBrowser
frmB.StartingAddress = "http://members.lycos.nl/tatlicocuk21/registreren/registreer.htm"
frmB.Show
End Sub

Private Sub nieuwefilm_Click()
End Sub

Private Sub printfilmlijst_Click()
'shel = ShellExecute(o&, vbNullString, "C:\Program Files\Mutlu'z Movie database\HTML1\filmlijst.html", vbNullString, vbNullString, SW_Shownormal)


MsgBox "N.V.T. in deze versie!"
End Sub

Private Sub selecteercover_Click()
End Sub

Private Sub skin1_Click()
FrmMain.Image13.Visible = False
FrmMain.Image12.Visible = True
FrmMain.Image12.Picture = LoadPicture(App.Path & "\black.jpg")
End Sub

Private Sub skin2_Click()
FrmMain.Image13.Visible = False
FrmMain.Image12.Visible = False
FrmMain.Picture14.Visible = True
End Sub

Private Sub skin3_Click()
End Sub

Private Sub sorteermenu_Click()
Dim f As New sorteer
            f.Show
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"

        Case "Open"
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
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Camera"
            MsgBox "N.V.T. in deze versie!"
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Back"
            Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
        Case "Forward"
            Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 1
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
'txtpost = Str(postnr)
visapost postnr
End If

        Case "Find"
            Dim f As New FrmFind
            f.Show
            
        Case "Sort Ascending"
            MsgBox "N.V.T. in deze versie!"
        Case "Sort Descending"
          MsgBox "N.V.T. in deze versie!"
          
    End Select
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
MsgBox "N.V.T. in deze versie!"
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "All Files (*.*)|*.*"
            .ShowSave
            If Len(.filename) = 0 Then
                Exit Sub
            End If
            sFile = .filename
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub



Private Sub totaaldatabasemmdb_Click()
End Sub

Private Sub uitgeleend_Click()

End Sub

Private Sub update_Click()
End Sub

Private Sub verwijder_Click()
End Sub

Private Sub verwijderalles_Click()
End Sub

Private Sub updatee_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://members.lycos.nl/tatlicocuk21/mmdbupdate/update.htm"
    frmB.Show
End Sub

Private Sub videoland_Click()
Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.videoland.nl"
    frmB.Show
End Sub

Private Sub Volgendefilm_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr + 1
If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
'txtpost = Str(postnr)
visapost postnr
End If
End Sub

Private Sub Vorigefilm_Click()
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
If Rs.RecordCount > 0 Then
postnr = postnr - 1
If postnr < 1 Then postnr = 1
txtpost = Str(postnr)
visapost postnr
End If
End Sub

Private Sub zoeken_Click()
Dim f As New FrmFind
            f.Show
'MsgBox "Vul de exacte filmnaam die je hebt opgegeven"

End Sub

Private Sub zoektrailer_Click()
MsgBox "N.V.T. in deze versie!"
'Dim frmB As New frmBrowser
'    frmB.StartingAddress = "http://members.home.nl/m.erol/search.htm"
'    frmB.Show
End Sub
