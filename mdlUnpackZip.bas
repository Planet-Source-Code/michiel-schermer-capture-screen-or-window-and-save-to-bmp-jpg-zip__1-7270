Attribute VB_Name = "mdlUnpackZip"
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

Private Declare Function windll_unzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPNAMES, ByVal xfnc As Long, ByRef xfnv As ZIPNAMES, lpDCL As DCLIST, lpUserFunc As USERFUNCTION) As Long
Private Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)
    Private sZipMsg As String
    Private lNumFilesInArchive As Long
    Private sZipInfo As String
    Private dblZIPUnpackedBytes As Double
    Private dblZIPPackedBytes As Double

    Private Type ZIPNAMES
        s(0 To 99) As String
    End Type
    
    Private Type CBCHAR
        ch(32800) As Byte
    End Type
    
    Private Type CBCH
        ch(256) As Byte
    End Type
    
    Private Type DCLIST
        ExtractOnlyNewer As Long
        SpaceToUnderscore As Long
        PromptToOverwrite As Long
        fQuiet As Long
        ncflag As Long
        ntflag As Long
        nvflag As Long
        nUflag As Long
        nzflag As Long
        ndflag As Long
        noflag As Long
        naflag As Long
        nZIflag As Long
        C_flag As Long
        fPrivilege As Long
        Zip As String
        ExtractDir As String
    End Type
    
    Private Type USERFUNCTION
        DllPrnt As Long
        DLLSND As Long
        DLLREPLACE As Long
        DLLPASSWORD As Long
        DllMessage As Long
        cchComment As Integer
        TotalSizeComp As Long
        TotalSize As Long
        CompFactor As Long
        NumMembers As Long
    End Type
    
    Private Type UZPVER
        structlen As Long
        flag As Long
        beta As String * 10
        date As String * 20
        zlib As String * 10
        unzip(1 To 4) As Byte
        zipinfo(1 To 4) As Byte
        os2dll As Long
        windll(1 To 4) As Byte
    End Type

Public Function UnpackZIP(sFileToUnzip As String, strExtractTo As String) As Boolean
    Dim ocolFiles As Collection
    Dim ocolXFiles As Collection
    
    UnpackZIP = VBUnzip(sFileToUnzip, strExtractTo, 0, 0, 0, 1, ocolFiles, ocolXFiles)
End Function

Private Function VBUnzip(sFile As String, sExtdir As String, nPrompOverWr As Integer, nAlwaysOverWr As Integer, nVerboseList As Integer, nArgsAreDirs As Integer, colFiles As Collection, colXFiles As Collection) As Boolean
    Dim MYDCL As DCLIST
    Dim MYUSER As USERFUNCTION
    Dim MYVER As UZPVER
    Dim lNumFiles As Long
    Dim vbzipnam As ZIPNAMES
    Dim lNumXFiles As Long
    Dim vbxnames As ZIPNAMES
    Dim X As Long
    Dim glRet As Long
    
    sZipInfo = ""
    lNumFilesInArchive = 0
    sZipMsg = ""
    vbzipnam.s(0) = vbNullString
    vbxnames.s(0) = vbNullString
    
    If Not colFiles Is Nothing Then
        lNumFiles = colFiles.Count
        For X = 1 To colFiles.Count
            vbzipnam.s(X) = colFiles.Item(X)
        Next X
    End If
    
    If Not colXFiles Is Nothing Then
        lNumXFiles = colXFiles.Count
        For X = 1 To colXFiles.Count
            vbxnames.s(X) = colXFiles.Item(X)
        Next X
    End If
    
    With MYDCL
        .ExtractOnlyNewer = 0
        .SpaceToUnderscore = 0
        .PromptToOverwrite = nPrompOverWr
        .fQuiet = 0
        .ncflag = 0
        .ntflag = 0
        .nvflag = nVerboseList
        .nUflag = 0
        .nzflag = 0
        .ndflag = nArgsAreDirs
        .noflag = nAlwaysOverWr
        .naflag = 0
        .nZIflag = 0
        .C_flag = 0
        .fPrivilege = 0
        .Zip = sFile
        .ExtractDir = sExtdir
    End With
    
    With MYUSER
        .DllPrnt = FnPtr(AddressOf CBDllPrnt)
        .DLLSND = 0&
        .DLLREPLACE = FnPtr(AddressOf CBDllRep)
        If nVerboseList = 1 Then
            .DllMessage = FnPtr(AddressOf CBDllCountFiles)
        End If
    End With
    
    With MYVER
        .structlen = Len(MYVER)
        .beta = Space$(9) & vbNullChar
        .date = Space$(19) & vbNullChar
        .zlib = Space$(9) & vbNullChar
    End With
    
    UzpVersion2 MYVER
    glRet = windll_unzip(lNumFiles, vbzipnam, lNumXFiles, vbxnames, MYDCL, MYUSER)
    
    If glRet <> 0 Then
        VBUnzip = False
    Else
        VBUnzip = True
    End If
End Function

Public Function FnPtr(ByVal lp As Long) As Long
    FnPtr = lp
End Function

Private Function CBDllPrnt(ByRef tInfo As CBCHAR, ByVal lChars As Long) As Long
  On Error Resume Next
    Dim sInfo As String
    Dim X As Long
    Dim nCPos As Long
    
        For X = 0 To lChars
            If tInfo.ch(X) = 0 Then Exit For
            sInfo = sInfo & Chr$(tInfo.ch(X))
        Next X
    sZipInfo = sZipInfo & sInfo
        If Asc(sInfo) <> 10 Then
            nCPos = InStr(1, sInfo, vbLf)
            If nCPos > 0 Then
                sInfo = Mid$(sInfo, nCPos + 1)
            End If
                
            DoEvents
        End If
    CBDllPrnt = 0
End Function

Private Function CBDllRep(ByRef tFName As CBCHAR) As Long
    On Error Resume Next
    
    Dim sFile As String
    Dim X As Long
    Dim lRet As Long
    
    CBDllRep = 100
    
    For X = 0 To 255
        If tFName.ch(X) = 0 Then Exit For
        sFile = sFile & Chr$(tFName.ch(X))
    Next X
    
    CBDllRep = 102
End Function

Private Sub CBDllMessage(ByVal lUnPackSize As Long, ByVal lPackSize As Long, ByVal nCompFactor As Integer, ByVal nMonth As Integer, ByVal nDay As Integer, ByVal nYear As Integer, ByVal nHour As Integer, ByVal nMinute As Integer, ByVal c As Byte, ByRef tFName As CBCH, ByRef tMethod As CBCH, ByVal lCRC As Long, ByVal fcrypt As Byte)
  On Error Resume Next
    Dim sBuff As String * 128
    Dim sFile As String
    Dim sMethod As String
    Dim X As Long
    
    sBuff = Space$(128)
    If lNumFilesInArchive = 0 Then
        Mid$(sBuff, 1, 50) = "Filename:"
        Mid$(sBuff, 53, 4) = "Size"
        Mid$(sBuff, 62, 4) = "Date"
        Mid$(sBuff, 71, 4) = "Time"
        sZipMsg = sBuff & vbCrLf
        sBuff = Space$(128)
    End If
    
    For X = 0 To 255
        If tFName.ch(X) = 0 Then Exit For
        sFile = sFile & Chr$(tFName.ch(X))
    Next X
    
    Mid$(sBuff, 1, 50) = Mid$(sFile, 1, 50)
    Mid$(sBuff, 51, 7) = right$(Space$(7) & CStr(lPackSize), 7)
    Mid$(sBuff, 60, 3) = right$(CStr(nDay), 2) & "."
    Mid$(sBuff, 63, 3) = right$("0" & CStr(nMonth), 2) & "."
    Mid$(sBuff, 66, 2) = right$("0" & CStr(nYear), 2)
    Mid$(sBuff, 70, 3) = right$(CStr(nHour), 2) & ":"
    Mid$(sBuff, 73, 2) = right$("0" & CStr(nMinute), 2)
    Mid$(sBuff, 76, 2) = right$(" " & CStr(nCompFactor), 2)
    Mid$(sBuff, 79, 8) = right$(Space$(8) & CStr(lUnPackSize), 8)
    Mid$(sBuff, 88, 8) = right$(Space$(8) & CStr(lCRC), 8)
    Mid$(sBuff, 97, 2) = Hex$(c)
    Mid$(sBuff, 100, 2) = Hex$(fcrypt)
    
    For X = 0 To 255
        If tMethod.ch(X) = 0 Then Exit For
        sMethod = sMethod & Chr$(tMethod.ch(X))
    Next X
    
    sZipMsg = sZipMsg & sBuff & vbCrLf
    sZipMsg = sZipMsg & sMethod & vbCrLf
    lNumFilesInArchive = lNumFilesInArchive + 1
End Sub

Private Sub CBDllCountFiles(ByVal lUnPackSize As Long, ByVal lPackSize As Long, ByVal nCompFactor As Integer, ByVal nMonth As Integer, ByVal nDay As Integer, ByVal nYear As Integer, ByVal nHour As Integer, ByVal nMinute As Integer, ByVal c As Byte, ByRef tFName As CBCH, ByRef tMethod As CBCH, ByVal lCRC As Long, ByVal fcrypt As Byte)
  On Error Resume Next
    Dim sFile As String
    Dim X As Long
    
    For X = 0 To 255
        If tFName.ch(X) = 0 Then Exit For
        sFile = sFile & Chr$(tFName.ch(X))
    Next X
    
    If right$(sFile, 1) <> "/" And lPackSize <> 0 And nCompFactor <> 0 And lUnPackSize <> 0 Then
        dblZIPUnpackedBytes = dblZIPUnpackedBytes + lUnPackSize
        dblZIPPackedBytes = dblZIPPackedBytes + lPackSize
        lNumFilesInArchive = lNumFilesInArchive + 1
    End If
End Sub

