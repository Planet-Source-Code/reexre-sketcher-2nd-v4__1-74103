Attribute VB_Name = "modBLT"
'Declarations
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE = vbPaletteModeNone    'You can find other modes in the "PaletteModeConstants" section of your Object Browser


Public Const KDegToRad = 1.74532925199433E-02

'Loads an image on your hard drive to DC memory
Public Function LoadPic(DestID As Long, Path As String) As Long
    DestID = CreateCompatibleDC(GetDC(0))
    LoadPic = SelectObject(DestID, LoadPicture(Path))
End Function

'Unloads an image stored in DC memory to prevent memory leaks
Public Function UnloadPic(SrcID As Long) As Long
    UnloadPic = DeleteDC(SrcID)
End Function

