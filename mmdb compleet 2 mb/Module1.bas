Attribute VB_Name = "Module1"
Public Db As Database
Public Rs As Recordset
Public findstart As Integer
Public postnr As Integer
Public g As Long
Public Sub CreateNewDB(filename As String)
Dim NewDB As Database
Dim NewTable As TableDef
Dim DBName As String
        
    DBName = App.Path & "\" & filename
        
    Close

    If Dir(DBName) <> "" Then
        Kill DBName
    End If
    

    Set NewDB = CreateDatabase(DBName, dbLangGeneral)
                
       
    Set NewTable = NewDB.CreateTableDef("filmlijst")
    
    With NewTable
        .Fields.Append .CreateField("filmnaam", dbMemo)
        .Fields.Append .CreateField("acteurs", dbMemo)
        .Fields.Append .CreateField("genre", dbMemo)
        .Fields.Append .CreateField("type", dbMemo)
        .Fields.Append .CreateField("speelduur", dbMemo)
        .Fields.Append .CreateField("jaar", dbMemo)
        .Fields.Append .CreateField("rating", dbMemo)
        .Fields.Append .CreateField("uitgeleend", dbMemo)
        .Fields.Append .CreateField("info", dbMemo)
        .Fields.Append .CreateField("date", dbMemo)
        .Fields.Append .CreateField("picture", dbMemo)
    '    .Fields.Append .CreateField("aantalcds", dbMemo)
    '    .Fields.Append .CreateField("ondertiteling", dbMemo)
        
        
        
        For t = 0 To 10
            .Fields(t).AllowZeroLength = True
        Next t
    End With
        
        NewDB.TableDefs.Append NewTable
    
    NewDB.Close
End Sub
Public Sub find(findstr As String, start As Integer)
Dim found As Boolean, t As Integer
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Rs.MoveFirst
For t = start To Rs.RecordCount - 1
    With Rs
        For f = 0 To .Fields.Count - 1
            If .Fields(f) <> "" Then
            test = InStr(1, findstr, .Fields(f), vbTextCompare)
            'MsgBox test
            If InStr(1, findstr, Trim(.Fields(f))) > 0 Then GoTo found
            End If
        Next f
        Rs.MoveNext
    End With
    
Next t

Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Exit Sub
found:
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
findstart = t
visapost t + 1
postnr = t + 1
End Sub

Public Sub visapost(post As Integer)
Set Db = OpenDatabase(App.Path & "\mutlu.mdb")
Set Rs = Db.OpenRecordset("filmlijst")
Dim filmlijst As String
    Rs.Move post - 1
    FrmMain.txtfilmnaam = Trim(Rs.Fields(0))
    FrmMain.txtacteurs = Trim(Rs.Fields(1))
    FrmMain.txtgenre = Trim(Rs.Fields(2))
    FrmMain.txttype = Trim(Rs.Fields(3))
    FrmMain.txtspeelduur = Trim(Rs.Fields(4))
    FrmMain.txtjaar = Trim(Rs.Fields(5))
    FrmMain.txtrating = Trim(Rs.Fields(6))
    FrmMain.txtuitgeleend = Trim(Rs.Fields(7))
    FrmMain.txtmisc = Trim(Rs.Fields(8))
    FrmMain.lblcreated.Caption = "Contact created " & Trim(Rs.Fields(9))
    FrmMain.txtcd = Trim(Rs.Fields(0))
  '  FrmMain.aantalcds = Trim(Rs.Fields(11))
  '  FrmMain.ondertiteling = Trim(Rs.Fields(12))
    
    
    
    
    If Rs.Fields(10) > "" Then
        filmlijst = Rs.Fields(10)
    
    Open App.Path & "\tmpfile" For Binary As #1
    Put 1, , filmlijst
    Close
    FrmMain.picpic.Picture = LoadPicture(App.Path & "\tmpfile")
    Kill App.Path & "\tmpfile"
    
    If FrmMain.picpic.ScaleWidth >= FrmMain.Picture2.ScaleWidth Then FrmMain.picpic.Left = 0
    If FrmMain.picpic.ScaleHeight >= FrmMain.Picture2.ScaleHeight Then FrmMain.picpic.Top = 0
    If FrmMain.picpic.ScaleWidth < FrmMain.Picture2.ScaleWidth Then
        FrmMain.picpic.Left = (FrmMain.Picture2.ScaleWidth / 2) - (FrmMain.picpic.ScaleWidth / 2)
    End If
    If FrmMain.picpic.ScaleHeight < FrmMain.Picture2.ScaleHeight Then
        FrmMain.picpic.Top = (FrmMain.Picture2.ScaleHeight / 2) - (FrmMain.picpic.ScaleHeight / 2)
    End If
    FrmMain.Picture2.Cls
    FrmMain.picpic.Visible = True
    Else
    FrmMain.picpic.Visible = False
    FrmMain.Picture2.Cls
    FrmMain.Picture2.Print
    FrmMain.Picture2.Print
    FrmMain.Picture2.Print
    FrmMain.Picture2.Print
    FrmMain.Picture2.Print
    FrmMain.Picture2.Print "           Geen plaatje beschikbaar."
    
    End If
    
    FrmMain.StatusBar1.Panels(1).Text = "Er zijn" & Str(Rs.RecordCount) & " Films in de database. Film nr:" & Str(post) & "."
    FrmMain.txtpost.Text = Str(post)
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
End Sub
