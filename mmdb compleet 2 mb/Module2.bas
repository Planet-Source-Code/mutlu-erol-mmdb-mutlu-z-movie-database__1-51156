Attribute VB_Name = "Modulereg1"
Public Sub DoG(FormName As Object)
On Error Resume Next
    Dim i As Integer, y As Integer
    FormName.AutoRedraw = True
    FormName.DrawStyle = 6
    FormName.DrawMode = 13
    FormName.DrawWidth = 13
    FormName.ScaleMode = 3
    FormName.ScaleHeight = 256
    For i = 0 To 510
        FormName.Line (0, y)-(FormName.Width, y + 1), RGB(0, 100, i), BF
        y = y + 1
    Next i
End Sub


