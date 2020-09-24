VERSION 5.00
Begin VB.Form FExample 
   Caption         =   " Example"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cscreen 
      Caption         =   "Capture Screen"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox Pscreen 
      Height          =   4680
      Left            =   120
      ScaleHeight     =   4620
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SaveImage(pbImage As PictureBox, sFile As String)
  pbImage.Picture = pbImage.Image
  SavePicture pbImage.Picture, sFile
End Sub

Private Sub Cscreen_Click()

    'TakeScreenShot Pscreen, App.Path & "\ScreenShot.bmp"
        
    'Neccessery otherwise the mouse pointer will not be shown up in the saved picture
    Pscreen.AutoRedraw = True
    Pscreen.AutoSize = True
    
    'Capture the screen
    Pscreen.Picture = CaptureScreen
    'Capture the mouse
    PaintCursor Pscreen
    
    Pscreen.Picture = Pscreen.Image
    
    'Finaly Save
    'SavePicture Pscreen.Picture, App.Path & "\ScreenShot.bmp"
    SaveImage Pscreen, App.Path & "\ScreenShot.bmp"
    
End Sub
