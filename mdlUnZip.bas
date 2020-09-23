Attribute VB_Name = "mdlUnZip"
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

    Private Type ZFHeader
       Signature      As Long
       version        As Integer
       GPBFlag        As Integer
       Compress       As Integer
       date           As Integer
       Time           As Integer
       CRC32          As Long
       CSize          As Long
       USize          As Long
       FNameLen       As Integer
       ExtraField     As Integer
    End Type
    
Public GetZipFilename As GetCompressFileName
    Private Type GetCompressFileName
        strZipFile(1 To 9999) As String
        Filenumber As Integer
    End Type


Public Function CheckZIPfiles(ByVal ZipFile As String) As String
'  On Error Resume Next
    Dim FNum As Integer
    Dim iCounter As Integer
    Dim sResult As String
    Dim zhdr As ZFHeader
    Dim i As Integer
    
    ReadyState = False
    
    Const ZIPSIG = &H4034B50
    FNum = FreeFile
    Close FNum
    Open ZipFile For Binary As #FNum
    Get #FNum, , zhdr
    While zhdr.Signature = ZIPSIG
    ReDim s(0 To zhdr.FNameLen - 1) As String * 1
        
    For iCounter = 0 To UBound(s)
        s(iCounter) = Chr$(0)
    Next
    
    For iCounter = 0 To zhdr.FNameLen - 1
        Get #FNum, , s(iCounter)
    Next
    
    Seek #FNum, Seek(FNum) + zhdr.CSize + zhdr.ExtraField
        
    sResult = ""
    For iCounter = 0 To UBound(s)
        sResult = sResult & s(iCounter)
    Next
    
    Get #FNum, , zhdr
    Wend
    Close FNum

    CheckZIPfiles = sResult
End Function

Public Sub ExtractZipFile(strFileToUnzip As String)
    Dim PictureName As String
    PictureName = CheckZIPfiles(strFileToUnzip)
End Sub
