Attribute VB_Name = "Procedure"
Global X As ListItem
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_Shownormal = 1

'+++++++++++++++++++++++++++++++++++++
'  INTEGER DECLARATION
'+++++++++++++++++++++++++++++++++++++

Global wiz_counter%
Global i%
Global j%
Global f%

'+++++++++++++++++++++++++++++++++++++
'  LONG DECLARATION
'+++++++++++++++++++++++++++++++++++++
Global counter As Long
Global shel As Long

'+++++++++++++++++++++++++++++++++++++
'  RECORDSET DECLARATION
'+++++++++++++++++++++++++++++++++++++

Global cn As New ADODB.Connection
Global Table As New ADODB.Recordset
Global Field As New ADODB.Recordset
Global SQL_Database As New ADODB.Recordset
Global HTML As New ADODB.Recordset

'+++++++++++++++++++++++++++++++++++++
'  BOOLEAN DECLARATION
'+++++++++++++++++++++++++++++++++++++

Global Proceed As Boolean
Global Table_Added As Boolean

'+++++++++++++++++++++++++++++++++++++
'  STRING DECLARATION
'+++++++++++++++++++++++++++++++++++++
Global FOLDER$
Global Converter$
Global File_name$
Global FOLDER_PATH$
Global FILE_PATH$
Global Database_Type$
Global Temp_Table$
Global Field_list$
Global SQL_Query$
Global Check_Database$
Global Virtual_Path$
Global Database_path$
Global arr() As String
Global Field_added$

'+++++++++++++++++++++++++++++++++++++
'  PROCEDURES DECLARATION
'+++++++++++++++++++++++++++++++++++++

Public Sub AccessTables()
On Error GoTo jump

  If Proceed = True Then
    If Table.State = 1 Then Table.Close
    Set Table = cn.OpenSchema(adSchemaTables) 'Open the table Names
        
   Convert.cbotable.Clear
   
   Convert.cbotable.AddItem Space(10) & "----- TABLE -----"
   Convert.cbotable.AddItem ""
    'Fill The User Table Names Not System Table "MSYS" Is system table
    While Not Table.EOF
      If UCase(Left(Table!Table_name, 4)) <> "MSYS" Then
       If Table!TABLE_TYPE = "TABLE" Then
       Convert.cbotable.AddItem Table!Table_name
       End If
      End If
    Table.MoveNext
    Wend
    
    
    
    Table.MoveFirst
    'Fill The QUERY
    Convert.cbotable.AddItem ""
    Convert.cbotable.AddItem Space(10) & "----- QUERY -----"
    Convert.cbotable.AddItem ""
    
    While Not Table.EOF
       If Table!TABLE_TYPE = "VIEW" Then
       Convert.cbotable.AddItem Table!Table_name
       End If
    Table.MoveNext
    Wend
  
    Proceed = True
  End If

Exit Sub
jump:
MsgBox Err.Description, vbCritical
Proceed = False
End Sub

Public Sub SQLTables()
On Error GoTo jump

  If Table.State = 1 Then Table.Close
  Table.Open "select name from sysobjects where xtype='U'", cn, adOpenDynamic, adLockOptimistic
  
  Convert.cbotable.Clear
  
  Convert.cbotable.AddItem Space(10) & "----- TABLE -----"
  Convert.cbotable.AddItem ""
  
  While Not Table.EOF   'Fill The User Table Names Not System Table "MSYS" Is system table
   If Table.Fields(0) <> "dtproperties" Then
     Convert.cbotable.AddItem Table.Fields(0)
   End If
  Table.MoveNext
  Wend
  
  'Fill The QUERY
  Convert.cbotable.AddItem ""
  Convert.cbotable.AddItem Space(10) & "----- QUERY -----"
  Convert.cbotable.AddItem ""
    
  Proceed = True
  
Exit Sub
jump:
MsgBox Err.Description, vbCritical
Proceed = False
End Sub

Public Sub SqlConnect()
On Error GoTo jump

  If cn.State = 1 Then cn.Close
  cn.ConnectionString = "provider=sqloledb;server=" & Trim(Convert.txtserver) & ";user id=" & Trim(Convert.txtuser) & ";password=" & Trim(Convert.txtpass) & ";database=" & Convert.cbosqldatabase.Text
  cn.Open
  cn.CursorLocation = adUseClient
  
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Public Function GetFolderName(convertIn As String) As String
On Error GoTo jump
 
   'DEFAULT FOLDER NAME
   FOLDER = convertIn + "1"

   'Find And Make New Name Of Folder Until Folder Exists
   While Dir(Convert.txtpath.Text & FOLDER, vbDirectory) <> ""
     counter = counter + 1
     FOLDER = convertIn & counter
   Wend
   
   GetFolderName = FOLDER

Exit Function
jump:
MsgBox Err.Description, vbCritical
End
End Function

Function CreateASP()
On Error GoTo jump
 
 File_name = arr(i) & ".asp"
 FILE_PATH = FOLDER_PATH & "\" & File_name
 mhandle = FreeFile

 Open FILE_PATH For Output As mhandle

    Print #mhandle, "<%@ Language=VBScript %>"
    Print #mhandle, "<!--#include file=adovbs.inc-->"
    Print #mhandle, "<%"
    Print #mhandle, "'-------------------------------------------------------"
    Print #mhandle, "'File Name :" & Space(2) & File_name & ""
    Print #mhandle, "'Date      :" & Space(2) & Date & ""
    Print #mhandle, "'Time      :" & Space(2) & Time() & ""
    Print #mhandle, "'Designer  :" & Space(2) & "Deepak Sharma"
    Print #mhandle, "'Developer :" & Space(2) & "Deepak Sharma"
    Print #mhandle, "'E-Mail    :" & Space(2) & "Deepakmailto@rediffmail.com"
    Print #mhandle, "'For Any Suggestion Or Problem Please Write To on e-mail"
    Print #mhandle, "'--------------------------------------------------------"
    Print #mhandle, "%>"
    Print #mhandle, "<HTML>"
    Print #mhandle, "<HEAD>"
    Print #mhandle, "<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio6"">"
    Print #mhandle, "</HEAD>"
    Print #mhandle, "<BODY bgcolor=white>"
    Print #mhandle, "<center>"
    Print #mhandle, "<font color=black size=5><u>" & arr(i) & "</u></font>"
    Print #mhandle, "&nbsp;&nbsp;&nbsp;<a href=main.html>Main Menu</a>"
    Print #mhandle, "</center>"
    Print #mhandle, "<form method=post>"
    Print #mhandle, "<%"
    Print #mhandle, "Set db = server.CreateObject(""ADODB.Connection"")"
    Print #mhandle, "Set rs = server.CreateObject(""ADODB.Recordset"")"
    Print #mhandle, "db.Open """ & Database_path & """"
    Print #mhandle, "db.CursorLocation = 3 'client side cursor"
    Print #mhandle, "rs.Open """ & SQL_Query & """" & ",db"
    Print #mhandle, "rs.PageSize = 8"
    Print #mhandle, "If Request(""PageNo"") = """" Then"
    Print #mhandle, "     PageNo = 1"
    Print #mhandle, "Else"
    Print #mhandle, "     PageNo = Request(""PageNo"")"
    Print #mhandle, "End If"
    Print #mhandle, ""
    Print #mhandle, "Mv = Request.Form(""scroll"")"
    Print #mhandle, ""
    Print #mhandle, "If Mv = ""prev"" Or Mv = ""next"" Then"
    Print #mhandle, "   Select Case Mv"
    Print #mhandle, "       Case ""prev"""
    Print #mhandle, "          If PageNo <> 1 Then"
    Print #mhandle, "             PageNo = PageNo - 1"
    Print #mhandle, "          End If"
    Print #mhandle, "       Case ""next"""
    Print #mhandle, "          If rs.AbsolutePage < rs.PageCount Then"
    Print #mhandle, "             PageNo = PageNo + 1"
    Print #mhandle, "          End If"
    Print #mhandle, "   End Select"
    Print #mhandle, "End If"
    Print #mhandle, ""
    Print #mhandle, "rs.AbsolutePage = PageNo"
    Print #mhandle, "'---------------------------"
    Print #mhandle, "'Print The Heading Of Table"
    Print #mhandle, "'---------------------------"
    Print #mhandle, "With response"
    Print #mhandle, ".write ""<table align=center bgcolor=ivory border=0 cellpadding=4 cellspacing=4>"""
    Print #mhandle, ".write ""<tr bgcolor=#cccccc>"""
    Print #mhandle, "   For f = 0 To rs.Fields.Count - 1"
    Print #mhandle, "     .write ""<th>"" & rs.Fields(f).Name "
    Print #mhandle, "   Next"
    Print #mhandle, ".write ""</tr>"""
    Print #mhandle, "'------------------------"
    Print #mhandle, "'Print The Data Of Table"
    Print #mhandle, "'------------------------"
    Print #mhandle, "While Not rs.EOF And Row < rs.PageSize"
    Print #mhandle, "  .write ""<tr bgcolor=#eeeeee>"""
    Print #mhandle, "  For f= 0 To rs.Fields.Count - 1"
    Print #mhandle, "    If IsNull(rs.Fields(f)) Then"
    Print #mhandle, "      .write ""<td>&nbsp;</td>"""
    Print #mhandle, "    Else"
    Print #mhandle, "      .write ""<td>"" & rs.Fields(f) & ""</td>"""
    Print #mhandle, "    End If"
    Print #mhandle, "  Next"
    Print #mhandle, "  .write ""</tr>"""
    Print #mhandle, "   row=row+1"
    Print #mhandle, "   rs.MoveNext"
    Print #mhandle, "Wend"
    'Print #mhandle, ".write ""</table>"""
    Print #mhandle, "End With"
    Print #mhandle, "%>"
    Print #mhandle, "<center>"
    Print #mhandle, "<input type=""hidden"" name=PageNo Value=""<%= PageNo %>"">"
    Print #mhandle, "<tr  bgcolor=#cccccc align=center>"
    Print #mhandle, "<td colspan=<%=rs.Fields.Count%>>"
    Print #mhandle, "<%if pageno>1 then%>"
    Print #mhandle, "<input type=Submit name=scroll value=prev>"
    Print #mhandle, "<%end if%>"
    Print #mhandle, "&nbsp;"
    Print #mhandle, "<%if pageno<rs.PageCount then%>"
    Print #mhandle, "<input type=Submit name=scroll value=next>"
    Print #mhandle, "<%end if%>"
    Print #mhandle, "</td>"
    Print #mhandle, "</tr>"
    Print #mhandle, "</table>"
    Print #mhandle, "</center>"
    Print #mhandle, "</form>"
    Print #mhandle, "</BODY>"
    Print #mhandle, "</HTML>"

 Close mhandle

Exit Function
jump:
MsgBox Err.Description, vbCritical
End
End Function

Function CreateHTML()
On Error GoTo jump

  If HTML.State = 1 Then HTML.Close
  HTML.Open SQL_Query, cn, adOpenDynamic, adLockOptimistic
  mhandle = FreeFile
  
  File_name = arr(i) & ".html"
  FILE_PATH = FOLDER_PATH & "\" & File_name
  mhandle = FreeFile
  
  Open FILE_PATH For Output As mhandle

    Print #mhandle, "<comment>"
    Print #mhandle, "-------------------------------------------------------"
    Print #mhandle, "File Name :" & Space(2) & File_name & ""
    Print #mhandle, "Date      :" & Space(2) & Date & ""
    Print #mhandle, "Time      :" & Space(2) & Time() & ""
    Print #mhandle, "Designer  :" & Space(2) & "Deepak Sharma"
    Print #mhandle, "Developer :" & Space(2) & "Deepak Sharma"
    Print #mhandle, "E-Mail    :" & Space(2) & "Deepakmailto@rediffmail.com"
    Print #mhandle, "For Any Suggestion Or Problem Please Write To on e-mail"
    Print #mhandle, "--------------------------------------------------------"
    Print #mhandle, "</comment>"
    Print #mhandle, "<HTML>"
    Print #mhandle, "<HEAD>"
    Print #mhandle, "<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">"
    Print #mhandle, "<TITLE></TITLE>"
    Print #mhandle, "</HEAD>"
    Print #mhandle, "<BODY bgcolor=white>"
    Print #mhandle, "<center>"
    Print #mhandle, "<font color=black size=5><u>" & arr(i) & "</u></font>"
    Print #mhandle, "&nbsp;&nbsp;&nbsp;"
    Print #mhandle, "<a href=main.html>Main Menu</a>"
    Print #mhandle, "</center>"
    Print #mhandle, "<table align=center bgcolor=ivory border=0 cellpadding=4 cellspacing=4>"
    Print #mhandle, "<tr bgcolor=#cccccc>"
    For f = 0 To HTML.Fields.Count - 1
      Print #mhandle, "<th>" & HTML.Fields(f).Name
    Next
      Print #mhandle, "</tr>"
     '------------------------
     'Print The Data Of Table"
     '------------------------
    Screen.MousePointer = vbHourglass
    While Not HTML.EOF
       Print #mhandle, "<tr bgcolor=#eeeeee>"
       For f = 0 To HTML.Fields.Count - 1
          If IsNull(HTML.Fields(f)) Then
             Print #mhandle, "<td>&nbsp;</td>"
          Else
             Print #mhandle, "<td>" & HTML.Fields(f) & "</td>"
          End If
       Next
       Print #mhandle, "</tr>"
       HTML.MoveNext
    Wend
    Print #mhandle, "</table>"
    Print #mhandle, "<center>"
    Print #mhandle, "<a href=" & File_name & ">Up</a>"
    Print #mhandle, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Print #mhandle, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Print #mhandle, "<a href=main.html>Main Menu</a>"
    Print #mhandle, "</center>"
    Screen.MousePointer = vbArrow
  
  Close mhandle
  
Exit Function
jump:
MsgBox Err.Description, vbCritical
End
End Function

Public Sub ASPConverter()
On Error GoTo jump
 
 Converter = "ASP"
 GetFields
 Call MainMenu(".asp")  'generate the main menu
 
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Public Sub HTMLConverter()
On Error GoTo jump

  Converter = "HTML"
  GetFields
  Call MainMenu(".html")  'generate the main menu
   
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Public Sub GetFields()
On Error GoTo jump
  
  Convert.DistinctName 'fill table name
   Field_list = ""
   SQL_Query = ""
   j = 0
   
  'FIND THAT THE PATH HAS \ AT LAST IF NOT THEN PUT \
   If Right(Convert.txtpath.Text, 1) <> "\" Then
      Convert.txtpath.Text = Convert.txtpath.Text + "\"
   End If
 
   FOLDER_PATH = Convert.txtpath.Text & GetFolderName(Converter)
   MkDir FOLDER_PATH
 
   With Convert.lstaddedfields
   
   For i = 1 To UBound(arr()) - 1
     For j = 1 To .ListItems.Count
        Set X = .ListItems.Item(j)
        If arr(i) = .ListItems.Item(j) Then
           Field_list = Field_list & "[" & X.SubItems(1) & "]" & ","
        End If
     Next
     
     'DELETE THE , FROM THE LAST
     Field_list = Mid(Field_list, 1, InStrRev(Field_list, ",", -1, vbTextCompare) - 1)
     SQL_Query = "Select " & Field_list & " from " & "[" & arr(i) & "]"
     
     'Identify the converter
     If Converter = "ASP" Then
       CreateASP 'CREATE THE ASP PAGE
     ElseIf Converter = "HTML" Then
       CreateHTML 'CREATE HTML PAGE
     End If
     
     Field_list = ""
     SQL_Query = ""
     
    Next
   End With
   
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub

Public Sub MainMenu(ext As String)
On Error GoTo jump

  mhandle = FreeFile
  Open FOLDER_PATH & "\Main.html" For Output As mhandle
  
    Print #mhandle, "<HTML>"
    Print #mhandle, "<HEAD>"
    Print #mhandle, "<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">"
    Print #mhandle, "<TITLE></TITLE>"
    Print #mhandle, "</HEAD>"
    Print #mhandle, "<BODY bgcolor=ivory>"
    Print #mhandle, "<br><br>"
    Print #mhandle, "<table align=center bgcolor=#0099FF border=1 cellpadding=5 cellspacing=4 width=""30%"">"
    Print #mhandle, "<td align=center><font face=arial size=5 color=#660000><b><u>Main Menu</u></b></font></td>"
    '--------------------------------------------------
    ' THIS LINE WILL FIND IF CONVERTOR IS ASP THEN SET
    ' THE VIRTUAL PATH ELSE NORMAL PATH
    '--------------------------------------------------
    For i = 1 To UBound(arr()) - 1
      Print #mhandle, "<tr bgcolor=#FFFFCC>"
      If ext = ".asp" Then
        Virtual_Path = "http:\\localhost\" & Mid(FOLDER_PATH, InStr(1, FOLDER_PATH, ":") + 2) & "\"
        Print #mhandle, "<td align=center><b><a href=" & Virtual_Path & arr(i) & ext & ">" & StrConv(arr(i), vbUpperCase) & "</a></b></td>"
        Virtual_Path = ""
      ElseIf ext = ".html" Then
        Print #mhandle, "<td align=center><b><a href=" & arr(i) & ext & ">" & StrConv(arr(i), vbUpperCase) & "</a></b></td>"
      End If
      Print #mhandle, "</tr>"
    Next
    Print #mhandle, "</table>"
    Print #mhandle, "</BODY>"
    Print #mhandle, "</HTML>"
    
  Close mhandle
  
Exit Sub
jump:
MsgBox Err.Description, vbCritical
End
End Sub
