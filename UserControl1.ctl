VERSION 5.00
Begin VB.UserControl CaptureScreen 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   450
   Windowless      =   -1  'True
   Begin VB.Image Image1 
      Height          =   420
      Left            =   15
      Picture         =   "UserControl1.ctx":0000
      Top             =   15
      Width           =   420
   End
End
Attribute VB_Name = "CaptureScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This OCX control is made by M. Schermer from the Netherlands
' This OCX can be used to capture the fullscreen or the active
' window. I'm take no care about damage on your computer but
' that is impossible.
'
' To use this control compile it to an OCX and start an new project
' Goto COMPONENTS and add the compiled OCX file
' Now you can use it to add the control to your form and for
' example:

'   Dim FileToOpen as String
'
'   FileToOpen = CaptureScreen1.CaptureActiveScreen("C:\MyScreen.BMP", BMP, True)
'   Image1.Picture = LoadPictures(FileToOpen)
'
' PS: If you want to make a loop (for example: refresh picture every 1 second)
' check the CaptureScreen1.ReadyState
' it returns True if the process is ready
' it returns False if the process is not ready
'
'                           Have fun using this control


    Public Enum WhatScreen
        FullScreen = 1
        ActiveWindow = 2
        FromWindowHwnd = 3
    End Enum

    Public Enum SaveInFormat
        BMP = 1
        JPG = 2
    End Enum

    Public Enum OpenInFormat
        BMPformat = 1
        JPGformat = 2
        ZIPformat = 3
    End Enum


Public Function CaptureFullScreen(strFileNameToSave As String, WhichFormat As SaveInFormat, UseZipCompression As Boolean) As String
    ReadyState = False
    Set m_cZ = New cZip
    Call CaptureAScreen(FullScreen, strFileNameToSave, WhichFormat, UseZipCompression)
    CaptureFullScreen = strOutputFile
End Function

Public Function CaptureActiveScreen(strFileNameToSave As String, WhichFormat As SaveInFormat, UseZipCompression As Boolean) As String
    ReadyState = False
    Set m_cZ = New cZip
    Call CaptureAScreen(ActiveWindow, strFileNameToSave, WhichFormat, UseZipCompression)
    CaptureActiveScreen = strOutputFile
End Function

Public Function LoadPictures(strFileNameToOpen As String) As Picture
  On Error GoTo ErrorHandler
    Dim UnpackResult As String
    Dim PictureName As String

    Set LoadPictures = LoadPicture(strFileNameToOpen) ' BMP format
    ReadyState = True
Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 481: GoTo UnpackZipFile
    End Select
Exit Function

UnpackZipFile:
    PictureName = CheckZIPfiles(strFileNameToOpen)
    If PictureName <> "" Then
        UnpackResult = UnpackZIP(strFileNameToOpen, "C:\")
        If UnpackResult = True Then
            Set LoadPictures = LoadPicture("C:\" & PictureName)
            Close
            Kill "C:\" & PictureName
            Kill strFileNameToOpen
            ReadyState = True
        End If
    End If
Exit Function

End Function

Public Function CheckState() As Boolean
    CheckState = ReadyState
End Function

Private Sub UserControl_Resize()
    UserControl.Height = 450
    UserControl.Width = 450
End Sub

Private Sub UserControl_Show()
    ReadyState = True
End Sub

Private Sub UserControl_Terminate()
    Close
End Sub
