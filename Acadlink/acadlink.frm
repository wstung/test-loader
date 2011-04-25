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
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1455
      Left            =   1200
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OLE OLE1 
      Class           =   "AutoCADLT.Drawing.18"
      Height          =   10770
      Left            =   240
      OleObjectBlob   =   "acadlink.frx":0000
      SizeMode        =   2  'AutoSize
      TabIndex        =   1
      Top             =   480
      Width           =   21300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sImageFileName As String
Private imageWidth As Long
Private imageHeight As Long

Private Sub Form_Load()
    sImageFileName = ""
End Sub

Private Sub Form_Resize()
   'Form1.OLE1.Height = Form1.ScaleHeight
   'Form1.OLE1.Width = Form1.ScaleWidth
   'Form1.OLE1.Left = Form1.ScaleLeft
   'Form1.OLE1.Top = Form1.ScaleTop
   'Form1.Picture1.Height = Form1.ScaleHeight
   'Form1.Picture1.Width = Form1.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    'Form1.OLE1.Close
End Sub

Private Sub OLE1_Updated(Code As Integer)
    On Error GoTo errorHandler
    If Code = 0 Then
    
        imageWidth = OLE1.Picture.width
        imageHeight = OLE1.Picture.height
        
        If sImageFileName <> "" Then
            Picture1.AutoRedraw = True
            Picture1.Picture = LoadPicture("")
            Picture1.width = OLE1.Picture.width
            Picture1.height = OLE1.Picture.height
            Picture1.PaintPicture OLE1.Picture, 0, 0, Picture1.width, Picture1.height, 0, 0, ScaleX(OLE1.Picture.width, vbHimetric, vbTwips), ScaleY(OLE1.Picture.height, vbHimetric, vbTwips)
            'Picture1.PaintPicture OLE1.Picture, 0, 0
            Set Picture1.Picture = Picture1.Image
        End If
    ElseIf Code = 2 Then
        savePicture Picture1.Picture, sImageFileName
        Unload Form1
    End If
    Exit Sub
    
errorHandler:
    MsgBox "[Ole Updated] Interface AutoCAD error: " + Err.Description
    Err.Clear
End Sub

Public Sub PopUp(width As Long, height As Long)
    Me.Show vbModal
    width = imageWidth
    height = imageHeight
End Sub

