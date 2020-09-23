Attribute VB_Name = "Module1"
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

    Public ReadyState As Boolean
    Public strOutputFile As String
    Private RectActive As Rect

Public Sub CaptureAScreen(ScreenToDo As WhatScreen, Optional strFileNameToSave As String, Optional WhichFormat As SaveInFormat, Optional UseZipCompression As Boolean, Optional HwndOfWindow As Long)
    Dim OriginalName As String
    
    OriginalName = strFileNameToSave
    Call CreateDirectory(strFileNameToSave)
    strFileNameToSave = OriginalName

    Select Case ScreenToDo
        Case 1: Form1.Picture = CaptureScreen
        Case 2: Form1.Picture = CaptureActiveWindow
        Case Else:
    End Select

    If strFileNameToSave <> "" Then
        Select Case WhichFormat
            Case 1: Call SavePictureToBMP(strFileNameToSave)
            Case 2: Call SavePictureToJPG(strFileNameToSave)
        End Select
    End If

    If strFileNameToSave <> "" Then
        Select Case UseZipCompression
            Case False: strOutputFile = strFileNameToSave
            Case True: Call CompressPicture(strFileNameToSave)
        End Select
    End If

End Sub

Public Function CaptureScreen() As Picture
    Set CaptureScreen = CaptureWindow(GetDesktopWindow, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

Public Function CaptureActiveWindow() As Picture
    Call GetWindowRect(GetForegroundWindow, RectActive)
    Set CaptureActiveWindow = CaptureWindow(GetForegroundWindow, False, 0, 0, RectActive.right - RectActive.left, RectActive.bottom - RectActive.top)
End Function

Public Function CaptureWindowHWND(WindowHWND As Long) As Picture
    Call GetWindowRect(WindowHWND, RectActive)
    Set CaptureWindowHWND = CaptureWindow(WindowHWND, False, 0, 0, RectActive.right - RectActive.left, RectActive.bottom - RectActive.top)
End Function


Public Sub SavePictureToBMP(strFileName As String)
    Dim CheckFile As Boolean
        
        CheckFile = FileExists(strFileName)
        If CheckFile = True Then
            Kill strFileName
        End If
    
    DoEvents
    Call SavePicture(Form1.Picture, strFileName)
    DoEvents
    strOutputFile = strFileName
End Sub

Public Sub SavePictureToJPG(strFileName As String)
    Dim c As New cDIBSection
    Dim CheckFile As Boolean
    Set c = New cDIBSection
    
        CheckFile = FileExists(strFileName)
        If CheckFile = True Then
            Kill strFileName
        End If
    
    DoEvents
    Call SavePicture(Form1.Picture, strFileName)
    DoEvents
    c.CreateFromPicture LoadPicture(strFileName)
    
    DoEvents
    Call SaveJPG(c, strFileName)
    DoEvents
    strOutputFile = strFileName

End Sub

Public Sub CompressPicture(strFileToCompress As String)
  On Error Resume Next
    Dim ZipFileName As String
    Dim CheckFile As Boolean
    
    ZipFileName = MakeZipName(strFileToCompress)
    
    Set m_cZ = New cZip
    
        CheckFile = FileExists(ZipFileName)
        If CheckFile = True Then
            Kill ZipFileName
        End If
    
    With m_cZ
        .ZipFile = ZipFileName
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFileToCompress
        .Zip
    End With
    
    Kill strFileToCompress
    strOutputFile = ZipFileName
End Sub

Private Function MakeZipName(strNameOfFile As String) As String
    Dim GotName As Boolean
    Dim FindChar As String * 1
    Dim FindPos As Integer
    
    Do Until GotName = True
        FindPos = FindPos + 1
        FindChar = right(strNameOfFile, FindPos)
            If FindPos <= Len(strNameOfFile) Then
                If FindChar = "." Then
                    MakeZipName = Mid(strNameOfFile, 1, Len(strNameOfFile) - (FindPos - 1)) & "zip"
                    GotName = True
                End If
            Else
                MakeZipName = strNameOfFile & ".zip"
                GotName = True
            End If
        DoEvents
    Loop
End Function


Private Function FileExists(filename) As Boolean
  On Error GoTo ErrorHandler
   FileExists = (Dir(filename) <> "")
Exit Function

ErrorHandler:
    FileExists = False
End Function

Private Sub CreateDirectory(strDirToCreate As String)
  On Error Resume Next
    Dim FindBeginPos As Integer
    Dim FindEndPos As Integer
    Dim DirIsCreated As Boolean
    Dim CreatePath As String
    
    FindBeginPos = InStr(strDirToCreate, ":")
    If FindBeginPos <> 0 Then
        CreatePath = Mid(strDirToCreate, FindBeginPos - 1, FindBeginPos) & "\"
            Do Until DirIsCreated = True
                FindBeginPos = InStr(strDirToCreate, "\")
                    If FindBeginPos <> 0 Then
                        FindEndPos = InStr(Mid(strDirToCreate, FindBeginPos + 1), "\")
                            If FindEndPos <> 0 Then
                                FindEndPos = (FindEndPos + FindBeginPos) - 2
                                CreatePath = CreatePath & Mid(strDirToCreate, FindBeginPos + 1, FindEndPos - 2) & "\"
                                MkDir CreatePath
                                strDirToCreate = Mid(strDirToCreate, FindEndPos)
                            Else
                                DirIsCreated = True
                            End If
                    End If
            Loop
    End If
End Sub

