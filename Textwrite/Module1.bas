Attribute VB_Name = "textwrite"
Declare Function CreateCaret Lib _
"user32" (ByVal hwnd As Long, _
ByVal hBitmap As Long, ByVal nWidth _
As Long, ByVal nHeight As Long) As Long

Declare Function ShowCaret Lib _
"user32" (ByVal hwnd As Long) As Long

Declare Function GetFocus Lib _
"user32" () As Long


