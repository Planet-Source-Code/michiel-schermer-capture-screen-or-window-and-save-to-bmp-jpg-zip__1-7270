VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8535
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   45
      Top             =   165
   End
   Begin CaptureScreens.CaptureScreen CaptureScreen2 
      Left            =   1575
      Top             =   135
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin CaptureScreens.CaptureScreen CaptureScreen1 
      Left            =   1065
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   2055
      Picture         =   "Form2.frx":0000
      Top             =   150
      Width           =   195
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7455
      Left            =   240
      Stretch         =   -1  'True
      Top             =   735
      Width           =   10740
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    ReadyState = True
End Sub

Private Sub Timer1_Timer()
Dim GetPicture As String

    If CaptureScreen1.CheckState = True Then
        GetPicture = CaptureScreen1.CaptureFullScreen("C:\1michiel.bmp", BMP, True)
        Image1.Picture = CaptureScreen2.LoadPictures(GetPicture)
        DoEvents
    End If
        DoEvents
End Sub
