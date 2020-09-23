VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form internetconnect 
   AutoRedraw      =   -1  'True
   Caption         =   "Super functie M.E."
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "internetconnect.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10320
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Zoek film"
      Height          =   255
      Left            =   12240
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   10320
      TabIndex        =   11
      Text            =   "Title?0187738"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ok"
      Height          =   255
      Left            =   14160
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10320
      TabIndex        =   9
      Top             =   2280
      Width           =   3375
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   10320
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1800
      Width           =   3795
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10095
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   17806
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"internetconnect.frx":09CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Filmnaam"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Search the Web page:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   6600
      TabIndex        =   6
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   90
   End
End
Attribute VB_Name = "internetconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim webdoc As HTMLDocument
'Dim texbody As HTMLBody
'Dim Texob As IHTMLTxtRange
Dim j As Integer

Private Sub cboAddress_Change()
cboAddress = Text4
 If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    WebBrowser1.Navigate cboAddress.Text
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Texob.findText(Text3.Text) = True Then
  j = j + 1
  Texob.Select
  Texob.Collapse (False)
Else
  MsgBox Str$(j) + " found"
  Command1.Enabled = True
  Texob.Collapse True
End If
End Sub

Private Sub Command2_Click()
cboAddress = Text4
End Sub

Private Sub Command3_Click()
Set a = Text5

Text4 = "www.imdb.com/" + a
cboAddress = Text4

End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "C:\Documents and Settings\gangster\Bureaublad\videoland.htm"
Command1.Enabled = False
j = 0
End Sub

Private Sub List1_Click()

End Sub

Private Sub Text4_Change()
Set a = Text5

Text4 = "www.imdb.com/" + a


End Sub

Private Sub timTimer_Timer()
    If WebBrowser1.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = WebBrowser1.LocationName
    Else
        Me.Caption = "Working..."
    End If

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Set webdoc = WebBrowser1.Document
'Dim Acollection As IHTMLElementCollection
Set Acollection = webdoc.All.tags("a")
For i = 0 To Acollection.length - 22
  '  List1.AddItem Acollection.Item(I).toString
Next
Label1.Caption = "Number of links: " + Str(Acollection.length)
Set texbody = webdoc.body
Set Texob = texbody.createTextRange()
Text1.Text = Texob.Text
Texob.moveToElementText Acollection.Item(1)
Text2.Text = Texob.Text
Texob.Select




Command1.Enabled = True
End Sub

