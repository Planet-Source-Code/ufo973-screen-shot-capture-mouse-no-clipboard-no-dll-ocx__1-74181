Attribute VB_Name = "MScreen"
Option Explicit
'**************************************
'Windows API/Global Declarations for :Display Current Mouse Pointer Image
'**************************************
' Get the handle of the window the mouse is over
Private Declare Function WindowFromPoint Lib "USER32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
' Retrieves the handle of the current cursor
Private Declare Function GetCursor Lib "USER32" () As Long
' Gets the coordinates of the mouse pointer
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
' Gets the PID of the window specified
Private Declare Function GetWindowThreadProcessId Lib "USER32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
' Gets the PID of the current program
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
' This attaches our program to whichever thread "owns" the cursor at the moment
Private Declare Function AttachThreadInput Lib "USER32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
' The next function draws the cursor to picCursor
' Note: If you want to display it in an Image control, use the GetDc API call
Private Declare Function DrawIcon Lib "USER32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' The POINTAPI type hold the (X,Y) for GetCursorPos()
Private Type POINTAPI
x As Long
y As Long
End Type
' The following are used for keeping the window always on top. This is optional.
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_TOPMOST = -1
Private Const SWP_NOTOPMOST = -2
Const iconSize As Integer = 9

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "GDI32.DLL" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "GDI32.DLL" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal HDCSRC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function GetForegroundWindow Lib "USER32.DLL" () As Long
Public Declare Function SelectPalette Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Type PicBmp
    Size As Long
    Type As Long
    HBMP As Long
    HPal As Long
    Reserved As Long
End Type

Public Type PALETTEENTRY
    PERed As Byte
    PEGreen As Byte
    PEBlue As Byte
    PEFlags As Byte
End Type

Public Type LOGPALETTE
    PALVersion As Integer
    PALNumEntries As Integer
    PALPalEntry(255) As PALETTEENTRY
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

' Paints the cursor image to the picturebox
Public Function PaintCursor(PictureBox As PictureBox)
 Dim pt As POINTAPI
 Dim hWnd As Long
 Dim dwThreadID, dwCurrentThreadID As Long
 Dim hCursor
 Dim threadid
 Dim CurrentThreadID
 
 ' Get the position of the cursor
 GetCursorPos pt
 ' Then get the handle of the window the cursor is over
 hWnd = WindowFromPoint(pt.x, pt.y)
 
 ' Get the PID of the thread
 threadid = GetWindowThreadProcessId(hWnd, vbNull)
 
 ' Get the thread of our program
 CurrentThreadID = App.threadid
 
 ' If the cursor is "owned" by a thread other than ours, attach to that thread and get the cursor
 If CurrentThreadID <> threadid Then
AttachThreadInput CurrentThreadID, threadid, True
hCursor = GetCursor()
AttachThreadInput CurrentThreadID, threadid, False
 
 ' If the cursor is owned by our thread, use GetCursor() normally
 Else
hCursor = GetCursor()
 End If
 
 ' Use DrawIcon to draw the cursor to picCursor
 DrawIcon PictureBox.hDC, pt.x - iconSize, pt.y - iconSize, hCursor
End Function

Public Function CreateBitmapPicture(ByVal HBMP As Long, ByVal HPal As Long) As Picture
  
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    On Error Resume Next
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .HBMP = HBMP
        .HPal = HPal
    End With

    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
    Set CreateBitmapPicture = IPic
  
End Function

Public Function CaptureWindow(ByVal HWNDSrc As Long, ByVal Client As Boolean, ByVal LeftSRC As Long, ByVal TopSRC As Long, ByVal WidthSRC As Long, ByVal HeightSRC As Long) As Picture
  
    Dim HDCMemory As Long
    Dim HBMP As Long
    Dim HBMPPrev As Long
    Dim HDCSRC As Long
    Dim HPal As Long
    Dim HPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
  
    On Error Resume Next
  
    If Client Then
        HDCSRC = GetDC(HWNDSrc)
    Else
        HDCSRC = GetWindowDC(HWNDSrc)
    End If
    HDCMemory = CreateCompatibleDC(HDCSRC)
    HBMP = CreateCompatibleBitmap(HDCSRC, WidthSRC, HeightSRC)
    HBMPPrev = SelectObject(HDCMemory, HBMP)
    RasterCapsScrn = GetDeviceCaps(HDCSRC, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(HDCSRC, SIZEPALETTE)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.PALVersion = &H300
        LogPal.PALNumEntries = 256
        GetSystemPaletteEntries HDCSRC, 0, 256, LogPal.PALPalEntry(0)
        HPal = CreatePalette(LogPal)
        HPalPrev = SelectPalette(HDCMemory, HPal, 0)
        RealizePalette HDCMemory
    End If
    BitBlt HDCMemory, 0, 0, WidthSRC, HeightSRC, HDCSRC, LeftSRC, TopSRC, vbSrcCopy
    HBMP = SelectObject(HDCMemory, HBMPPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        HPal = SelectPalette(HDCMemory, HPalPrev, 0)
    End If
    DeleteDC HDCMemory
    ReleaseDC HWNDSrc, HDCSRC
    Set CaptureWindow = CreateBitmapPicture(HBMP, HPal)
   
End Function

Public Function CaptureScreen() As Picture
  
    Dim HWNDScreen As Long
  
    On Error Resume Next

    HWNDScreen = GetDesktopWindow()
    Set CaptureScreen = CaptureWindow(HWNDScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
  
End Function

Public Function TakeScreenShot(PictureBox As PictureBox, ByVal FileOutput As String)

    PictureBox.Picture = CaptureScreen
    SavePicture PictureBox.Picture, FileOutput

End Function
