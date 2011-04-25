VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13965
   ScaleHeight     =   8880
   ScaleWidth      =   13965
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.OLE OLE1 
      Class           =   "DWGTrueView.Drawing.18"
      Height          =   11115
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0000
      SizeMode        =   2  'AutoSize
      TabIndex        =   0
      Top             =   0
      Width           =   20385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sImageFileName As String


Private Sub Form_Resize()
   Form1.OLE1.Height = Form1.ScaleHeight
   Form1.OLE1.Width = Form1.ScaleWidth
   Form1.OLE1.Left = Form1.ScaleLeft
   Form1.OLE1.Top = Form1.ScaleTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Picture1 As PictureBox
    If sImageFileName <> "" Then
        Picture1.AutoRedraw = True
        Picture1.PaintPicture OLE1.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, ScaleX(OLE1.Picture.Width, vbHimetric, vbTwips), ScaleY(OLE1.Picture.Height, vbHimetric, vbTwips)
        Set Picture1.Picture = Picture1.Image
        savePicture Picture1.Picture, sImageFileName
    End If
    Form1.OLE1.Close
End Sub

Private Sub OLE1_Updated(Code As Integer)
    'Unload Form1
End Sub

