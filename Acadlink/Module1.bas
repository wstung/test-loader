Attribute VB_Name = "Module1"
Sub Main()
    Dim app As New acadApp
    Dim width As Long
    Dim height As Long
    Dim ret As Boolean
    app.savePicture "C:\Users\wstung\Documents\SEMI\MRP\SernKuo Project\test.bmp"
    ret = app.Activate("test", "C:\Users\wstung\Documents\SEMI\MRP\SernKuo Project\BFY-T525a.dwg", 2, width, height)
End Sub
