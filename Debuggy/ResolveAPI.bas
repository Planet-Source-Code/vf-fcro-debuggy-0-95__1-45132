Attribute VB_Name = "ResolveAPI"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Public PROCRETBRK As Long
'Public PROCRETTYPE As Long
Public PROCRETBRK As New Collection






