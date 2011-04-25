VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acadApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim docName As String
Private imageWidth As Long
Private imageHeight As Long

Private Sub Class_Initialize()
    Load Form1
End Sub

Public Sub GetImageSize(width As Integer, height As Integer)
    On Error GoTo errorHandling
    width = imageWidth
    height = imageHeight
    Exit Sub
errorHandling:
    MsgBox "[GetImageSize]Interface Autocad Error getting Image Siage: " + Err.Description
    Err.Clear
End Sub
Public Function Activate(sTitle As String, sDoc As String, sizeMode As Integer, width As Long, height As Long) As Boolean
    On Error GoTo errorHandling
    Activate = False
    docName = sDoc
    Form1.OLE1.sizeMode = sizeMode
    Form1.OLE1.Class = "AutoCADLT.Drawing"
    Form1.OLE1.CreateLink sDoc
    Form1.OLE1.DoVerb
    Form1.Caption = sTitle
    Form1.PopUp imageWidth, imageHeight
    width = imageWidth
    height = imageHeight
    Activate = True
    Exit Function
errorHandling:
    MsgBox "[Activate]Interface AutoCAD Error: " + Err.Description
    Err.Clear
End Function

Public Sub savePicture(sImageName As String)
    Form1.sImageFileName = sImageName
End Sub

Private Sub GetVerbList(ByRef frm As Form, ByRef OLE1 As Control)
    Dim intX As Integer, intNumVerbs As Integer
    Dim strVerbList As String

    ' Update verb list.
    With frm!OLE1
        .Action = 1
        intNumVerbs = .ObjectVerbsCount
        For intX = 0 To intNumVerbs - 1
            strVerbList = strVerbList & .ObjectVerbs(intX) & "; "
        Next intX
    End With

    ' Display verbs in message box.
    MsgBox Left(strVerbList, Len(strVerbList) - 2)
End Sub

