Attribute VB_Name = "mZip"
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


Private Declare Function ZpInit Lib "vbzip10.dll" (ByRef tUserFn As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "vbzip10.dll" (ByRef tOpts As ZPOPT) As Long ' Set Zip options
Private Declare Function ZpGetOptions Lib "vbzip10.dll" () As ZPOPT ' used to check encryption flag only
Private Declare Function ZpArchive Lib "vbzip10.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPNAMES) As Long ' Real zipping action
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function GetForegroundWindow Lib "USER32" () As Long
Public Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function ijlInit Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlRead Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlWrite Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlGetLibVersion Lib "ijl11.dll" () As Long
Private Declare Function ijlGetErrorString Lib "ijl11.dll" (ByVal code As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    Public Const RASTERCAPS As Long = 38
    Public Const RC_PALETTE As Long = &H100
    Public Const SIZEPALETTE As Long = 104
    Private Const OF_WRITE = &H1
    Private Const OF_SHARE_DENY_WRITE = &H20
    Private Const GENERIC_WRITE = &H40000000
    Private Const GENERIC_READ = &H80000000
    Private Const FILE_SHARE_WRITE = &H2
    Private Const CREATE_ALWAYS = 2
    Private Const FILE_BEGIN = 0
    Private Const SECTION_MAP_WRITE = &H2
    Private Const GMEM_DDESHARE = &H2000
    Private Const GMEM_DISCARDABLE = &H100
    Private Const GMEM_DISCARDED = &H4000
    Private Const GMEM_FIXED = &H0
    Private Const GMEM_INVALID_HANDLE = &H8000
    Private Const GMEM_LOCKCOUNT = &HFF
    Private Const GMEM_MODIFY = &H80
    Private Const GMEM_MOVEABLE = &H2
    Private Const GMEM_NOCOMPACT = &H10
    Private Const GMEM_NODISCARD = &H20
    Private Const GMEM_NOT_BANKED = &H1000
    Private Const GMEM_NOTIFY = &H4000
    Private Const GMEM_SHARE = &H2000
    Private Const GMEM_VALID_FLAGS = &H7F72
    Private Const GMEM_ZEROINIT = &H40
    Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
    Private Const MAX_PATH = 260
        Global Setting As Integer

    Private Enum IJLERR
        IJL_OK = 0
        IJL_INTERRUPT_OK = 1
        IJL_ROI_OK = 2
        IJL_EXCEPTION_DETECTED = -1
        IJL_INVALID_ENCODER = -2
        IJL_UNSUPPORTED_SUBSAMPLING = -3
        IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
        IJL_MEMORY_ERROR = -5
        IJL_BAD_HUFFMAN_TABLE = -6
        IJL_BAD_QUANT_TABLE = -7
        IJL_INVALID_JPEG_PROPERTIES = -8
        IJL_ERR_FILECLOSE = -9
        IJL_INVALID_FILENAME = -10
        IJL_ERROR_EOF = -11
        IJL_PROG_NOT_SUPPORTED = -12
        IJL_ERR_NOT_JPEG = -13
        IJL_ERR_COMP = -14
        IJL_ERR_SOF = -15
        IJL_ERR_DNL = -16
        IJL_ERR_NO_HUF = -17
        IJL_ERR_NO_QUAN = -18
        IJL_ERR_NO_FRAME = -19
        IJL_ERR_MULT_FRAME = -20
        IJL_ERR_DATA = -21
        IJL_ERR_NO_IMAGE = -22
        IJL_FILE_ERROR = -23
        IJL_INTERNAL_ERROR = -24
        IJL_BAD_RST_MARKER = -25
        IJL_THUMBNAIL_DIB_TOO_SMALL = -26
        IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
        IJL_RESERVED = -99
    End Enum
    
    Private Enum IJLIOTYPE
        IJL_SETUP = -1&
        IJL_JFILE_READPARAMS = 0&
        IJL_JBUFF_READPARAMS = 1&
        IJL_JFILE_READWHOLEIMAGE = 2&
        IJL_JBUFF_READWHOLEIMAGE = 3&
        IJL_JFILE_READHEADER = 4&
        IJL_JBUFF_READHEADER = 5&
        IJL_JFILE_READENTROPY = 6&
        IJL_JBUFF_READENTROPY = 7&
        IJL_JFILE_WRITEWHOLEIMAGE = 8&
        IJL_JBUFF_WRITEWHOLEIMAGE = 9&
        IJL_JFILE_WRITEHEADER = 10&
        IJL_JBUFF_WRITEHEADER = 11&
        IJL_JFILE_WRITEENTROPY = 12&
        IJL_JBUFF_WRITEENTROPY = 13&
        IJL_JFILE_READONEHALF = 14&
        IJL_JBUFF_READONEHALF = 15&
        IJL_JFILE_READONEQUARTER = 16&
        IJL_JBUFF_READONEQUARTER = 17&
        IJL_JFILE_READONEEIGHTH = 18&
        IJL_JBUFF_READONEEIGHTH = 19&
        IJL_JFILE_READTHUMBNAIL = 20&
        IJL_JBUFF_READTHUMBNAIL = 21&
    End Enum
    
    Private Type JPEG_CORE_PROPERTIES_VB ' Sadly, due to a limitation in VB (UDT variable count)
        UseJPEGPROPERTIES As Long                      '// default = 0
        DIBBytes As Long ';                  '// default = NULL 4
        DIBWidth As Long ';                  '// default = 0 8
        DIBHeight As Long ';                 '// default = 0 12
        DIBPadBytes As Long ';               '// default = 0 16
        DIBChannels As Long ';               '// default = 3 20
        DIBColor As Long ';                  '// default = IJL_BGR 24
        DIBSubsampling As Long  ';            '// default = IJL_NONE 28
        JPGFile As Long 'LPTSTR              JPGFile;                32   '// default = NULL
        JPGBytes As Long ';                  '// default = NULL 36
        JPGSizeBytes As Long ';              '// default = 0 40
        JPGWidth As Long ';                  '// default = 0 44
        JPGHeight As Long ';                 '// default = 0 48
        JPGChannels As Long ';               '// default = 3
        JPGColor As Long           ';                  '// default = IJL_YCBCR
        JPGSubsampling As Long  ';            '// default = IJL_411
        JPGThumbWidth As Long ' ;             '// default = 0
        JPGThumbHeight As Long ';            '// default = 0
        cconversion_reqd As Long ';          '// default = TRUE
        upsampling_reqd As Long ';           '// default = TRUE
        jquality As Long ';                  '// default = 75.  100 is my preferred quality setting.
        jprops(0 To 19999) As Byte
    End Type
    
    Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    
    Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
    End Type
    
    Type Rect
        left As Long
        top As Long
        right As Long
        bottom As Long
    End Type
    
    Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
    End Type
    
    Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(255) As PALETTEENTRY
    End Type
    
    Public Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(7) As Byte
    End Type
    
    Public Type PicBmp
        Size As Long
        Type As Long
        hBmp As Long
        hPal As Long
        Reserved As Long
    End Type
    
    Private Type ZIPNAMES         ' argv
        s(0 To 1023)   As String
    End Type
    
    Private Type CBCHAR           ' Callback large "string" (sic)
        ch(0 To 4096)  As Byte
    End Type
    
    Private Type CBCH             ' Callback small "string" (sic)
        ch(0 To 255)   As Byte
    End Type
    
    Private Type ZIPUSERFUNCTIONS ' Store the callback functions
        lPtrPrint      As Long    ' Pointer to application's print routine
        lptrPassword   As Long    ' Pointer to application's password routine.
        lptrComment    As Long
        lptrService    As Long    ' callback function designed to be used for allowing the app to process Windows messages, or cancelling the operation as well as giving option of progress. If this function returns non-zero, it will terminate what it is doing. It provides the app with the name of the archive member it has just processed, as well as the original size.
    End Type
    
    Public Type ZPOPT
        date           As String  ' US Date (8 Bytes Long) "12/31/98"?
        szRootDir      As String  ' Root Directory Pathname (Up To 256 Bytes Long)
        szTempDir      As String  ' Temp Directory Pathname (Up To 256 Bytes Long)
        fTemp          As Long    ' 1 If Temp dir Wanted, Else 0
        fSuffix        As Long    ' Include Suffixes (Not Yet Implemented!)
        fEncrypt       As Long    ' 1 If Encryption Wanted, Else 0
        fSystem        As Long    ' 1 To Include System/Hidden Files, Else 0
        fVolume        As Long    ' 1 If Storing Volume Label, Else 0
        fExtra         As Long    ' 1 If Excluding Extra Attributes, Else 0
        fNoDirEntries  As Long    ' 1 If Ignoring Directory Entries, Else 0
        fExcludeDate   As Long    ' 1 If Excluding Files Earlier Than Specified Date, Else 0
        fIncludeDate   As Long    ' 1 If Including Files Earlier Than Specified Date, Else 0
        fVerbose       As Long    ' 1 If Full Messages Wanted, Else 0
        fQuiet         As Long    ' 1 If Minimum Messages Wanted, Else 0
        fCRLF_LF       As Long    ' 1 If Translate CR/LF To LF, Else 0
        fLF_CRLF       As Long    ' 1 If Translate LF To CR/LF, Else 0
        fJunkDir       As Long    ' 1 If Junking Directory Names, Else 0
        fGrow          As Long    ' 1 If Allow Appending To Zip File, Else 0
        fForce         As Long    ' 1 If Making Entries Using DOS File Names, Else 0
        fMove          As Long    ' 1 If Deleting Files Added Or Updated, Else 0
        fDeleteEntries As Long    ' 1 If Files Passed Have To Be Deleted, Else 0
        fUpdate        As Long    ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
        fFreshen       As Long    ' 1 If Freshing Zip File-Overwrite Only, Else 0
        fJunkSFX       As Long    ' 1 If Junking SFX Prefix, Else 0
        fLatestTime    As Long    ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
        fComment       As Long    ' 1 If Putting Comment In Zip File, Else 0
        fOffsets       As Long    ' 1 If Updating Archive Offsets For SFX Files, Else 0
        fPrivilege     As Long    ' 1 If Not Saving Privileges, Else 0
        fEncryption    As Long    ' Read Only Property!!!
        fRecurse       As Long    ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
        fRepair        As Long    ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
        flevel         As Byte    ' Compression Level - 0 = Stored 6 = Default 9 = Max
    End Type

Private Function plAddressOf(ByVal lPtr As Long) As Long
    plAddressOf = lPtr ' VB Bug workaround fn
End Function

Public Function VBZip(cZipObject As cZip, tZPOPT As ZPOPT, sFileSpecs() As String, iFileCount As Long) As Long
    Dim tUser As ZIPUSERFUNCTIONS
    Dim lR As Long
    Dim i As Long
    Dim sZipFile As String
    Dim tZipName As ZIPNAMES
    
    m_bCancel = False
    Set m_cZip = cZipObject
        If Not Len(Trim$(m_cZip.BasePath)) = 0 Then
            ChDir m_cZip.BasePath
        End If
    tUser.lPtrPrint = plAddressOf(AddressOf ZipPrintCallback)      ' Set address of callback functions
    tUser.lptrPassword = plAddressOf(AddressOf ZipPasswordCallback)
    tUser.lptrComment = plAddressOf(AddressOf ZipCommentCallback)
    tUser.lptrService = plAddressOf(AddressOf ZipServiceCallback)  ' not coded yet :-)
    lR = ZpInit(tUser)
    lR = ZpSetOptions(tZPOPT) ' Set options
        For i = 1 To iFileCount   ' Go for it!
            tZipName.s(i - 1) = sFileSpecs(i)
            DoEvents
        Next i
    tZipName.s(i) = vbNullChar
    sZipFile = cZipObject.ZipFile
    lR = ZpArchive(iFileCount, sZipFile, tZipName) ' Dit neemt tijd in beslag
    Debug.Print lR
    VBZip = lR
End Function

Private Function ZipServiceCallback(ByRef mname As CBCHAR, ByVal X As Long) As Long
  On Error Resume Next                  'Always Put This In Callback Routines!
    Dim iPos As Long
    Dim sInfo As String
    Dim bCancel As Boolean
    
    If X > 1 And X < 32000 Then       ' Check we've got a message:
        ReDim b(0 To X) As Byte       ' If so, then get the readable portion of it:
        CopyMemory b(0), mname, X
        DoEvents
        sInfo = StrConv(b, vbUnicode) ' Convert to VB string:
        iPos = InStr(sInfo, vbNullChar)
            If iPos > 0 Then
                sInfo = left$(sInfo, iPos - 1)
            End If
        m_cZip.Service sInfo, bCancel
            If bCancel Then
                ZipServiceCallback = 1
            Else
                ZipServiceCallback = 0
            End If
    End If
End Function

Private Function ZipPrintCallback(ByRef fname As CBCHAR, ByVal X As Long) As Long
    On Error Resume Next
    Dim iPos As Long
    Dim sFile As String
    
    If X > 1 And X < 32000 Then        ' Check we've got a message:
        ReDim b(0 To X) As Byte        ' If so, then get the readable portion of it:
        CopyMemory b(0), fname, X
        sFile = StrConv(b, vbUnicode)  ' Convert to VB string:
            If iPos > 0 Then
                sFile = left$(sFile, iPos - 1)
                DoEvents
            End If
        ReplaceSection sFile, "/", "\" ' Fix up backslashes:
        m_cZip.ProgressReport sFile    ' Tell the caller about it
    End If
    
    ZipPrintCallback = 0
End Function

Private Function ZipCommentCallback(ByRef s1 As CBCHAR) As CBCHAR
 On Error Resume Next    ' always put this in callback routines!
   
   s1.ch(0) = vbNullString ' not supported always return \0
   ZipCommentCallback = s1
End Function

Private Function ZipPasswordCallback(ByRef pwd As CBCH, ByVal X As Long, ByRef s2 As CBCH, ByRef Name As CBCH) As Long
  On Error Resume Next
    Dim bCancel As Boolean
    Dim sPassword As String
    Dim b() As Byte
    Dim lSize As Long

    ZipPasswordCallback = 1                   ' The default:
        If m_bCancel Then
            Exit Function
        End If
    m_cZip.PasswordRequest sPassword, bCancel ' Ask for password:
    sPassword = Trim$(sPassword)
        If bCancel Or Len(sPassword) = 0 Then     ' Cancel out if no useful password:
            m_bCancel = True
            Exit Function
        End If
    lSize = Len(sPassword)                    ' Put password into return parameter:
        If lSize > 254 Then
            lSize = 254
        End If
    b = StrConv(sPassword, vbFromUnicode)
    CopyMemory pwd.ch(0), b(0), lSize
    ZipPasswordCallback = 0                   ' Ask UnZip to process it:
End Function

Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Long
    Dim iPos As Long
    Dim iLastPos As Long
    Dim ReadyProcess As Boolean
    
    iLastPos = 1
    ReadyProcess = False
    
    Do Until ReadyProcess = True
        iPos = InStr(sString, Chr(0))
        If iPos <> 0 Then
            sString = Mid(sString, 1, iPos - 1)
        Else
            ReadyProcess = True
        End If
    Loop
    
    Do
        iPos = InStr(iLastPos, sString, "/")
        If (iPos > 1) Then
            Mid$(sString, iPos, 1) = "\"
            iLastPos = iPos + 1
        End If
    Loop While Not (iPos = 0)
    
    ReplaceSection = iLastPos
End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    DoEvents
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    DoEvents
    Set CreateBitmapPicture = IPic
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
        If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context For client area
        Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context For entire window
        End If
    hDCMemory = CreateCompatibleDC(hDCSrc) ' Create a memory device context for the copy process
    DoEvents
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc) ' Create a bitmap and place it in the memory DC
    hBmpPrev = SelectObject(hDCMemory, hBmp) ' Get screen properties
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then ' Create a copy of the system palette
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            hPalPrev = SelectPalette(hDCMemory, hPal, 0) ' Select the new palette into the memoryDC and realize it
            r = RealizePalette(hDCMemory)
        End If
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy) ' Copy the on-screen image into the memory DC
    DoEvents
    hBmp = SelectObject(hDCMemory, hBmpPrev)
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then ' If the screen has a palette get back the palette that was selected in previously
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
    DoEvents
    r = DeleteDC(hDCMemory) ' Release the device context resources back to the system
    r = ReleaseDC(hWndSrc, hDCSrc) ' Call CreateBitmapPicture to create a picture object from the bitmap and palette handles. Then return the resulting picture object.
    DoEvents
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function SaveJPG(ByRef cDib As cDIBSection, ByVal sFile As String, Optional ByVal lQuality As Long = 90) As Boolean
    Dim tJ As JPEG_CORE_PROPERTIES_VB
    Dim bFile() As Byte
    Dim lPtr As Long
    Dim lR As Long
    Dim tFnd As WIN32_FIND_DATA
    Dim hFile As Long
    Dim bFileExisted As Boolean
    Dim lFileSize As Long
   
    hFile = -1
    
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
    
    bFileExisted = (FindFirstFile(sFile, tFnd) <> -1)
        If bFileExisted Then
        Kill sFile
        End If
    tJ.DIBWidth = cDib.Width
    tJ.DIBHeight = -cDib.Height
    tJ.DIBBytes = cDib.DIBSectionBitsPtr
    tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
    bFile = StrConv(sFile, vbFromUnicode)
    ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
    bFile(UBound(bFile)) = 0
    lPtr = VarPtr(bFile(0))
    DoEvents
    CopyMemory tJ.JPGFile, lPtr, 4
    DoEvents
    tJ.JPGWidth = cDib.Width
    tJ.JPGHeight = cDib.Height
    tJ.jquality = lQuality
    lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
    If lR = IJL_OK Then
    If bFileExisted Then
    hFile = lopen(sFile, OF_WRITE Or OF_SHARE_DENY_WRITE)
    If hFile = 0 Then
    Else
    SetFileTime hFile, tFnd.ftCreationTime, tFnd.ftLastAccessTime, tFnd.ftLastWriteTime
    lclose hFile
    SetFileAttributes sFile, tFnd.dwFileAttributes
    End If
    End If
    lFileSize = tJ.JPGSizeBytes - tJ.JPGBytes
    SaveJPG = True
    Else
    Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to save to JPG " & lR, vbExclamation
    End If
    ijlFree tJ
    Else
    Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to initialise the IJL library: " & lR
    End If
End Function


