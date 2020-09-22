Attribute VB_Name = "mod_Rotate"
Const Pi = 3.14159265358979
Public Const Trans = Pi / 180
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Double)

    'Rotate the image in a picture box.

    'pic1 is the picture box with the bitmap to rotate

    'pic2 is the picture box to receive the rotated bitmap

    'theta is the angle of rotation

    
    Dim c1x As Integer, c1y As Integer
    Dim c2x As Integer, c2y As Integer
    Dim a As Double
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim n As Integer, r As Integer
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = (n - 1) * 2
    pic1hDc = pic1.hdc
    pic2hDc = pic2.hdc


    For p2x = 0 To n / 2


        For p2y = 0 To n / 2
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0& = GetPixel(pic1hDc, c1x + p1x, c1y + p1y)
            c1& = GetPixel(pic1hDc, c1x - p1x, c1y - p1y)
            c2& = GetPixel(pic1hDc, c1x + p1y, c1y - p1x)
            c3& = GetPixel(pic1hDc, c1x - p1y, c1y + p1x)
            If c0& <> -1 Then SetPixel pic2hDc, c2x + p2x, c2y + p2y, c0&
            If c1& <> -1 Then SetPixel pic2hDc, c2x - p2x, c2y - p2y, c1&
            If c2& <> -1 Then SetPixel pic2hDc, c2x + p2y, c2y - p2x, c2&
            If c3& <> -1 Then SetPixel pic2hDc, c2x - p2y, c2y + p2x, c3&
        Next

        
        
    Next
End Sub
Sub bmp_rotate2(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Double)

    'Rotate the image in a picture box.

    'pic1 is the picture box with the bitmap to rotate

    'pic2 is the picture box to receive the rotated bitmap

    'theta is the angle of rotation

    
    Dim c1x As Double, c1y As Double
    Dim c2x As Double, c2y As Double
    Dim a As Double
    Dim p1x As Double, p1y As Double
    Dim p2x As Double, p2y As Double
    Dim n As Integer, r As Double
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = (n - 1) * 2
    pic1hDc = pic1.hdc
    pic2hDc = pic2.hdc


    For p2x = 0 To n / 2


        For p2y = 0 To n / 2
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x) 'Tan(p2y / p2x)  '
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0& = pPoint(c1x + p1x, c1y + p1y, pic1hDc)
            c1& = pPoint(c1x - p1x, c1y - p1y, pic1hDc)
            c2& = pPoint(c1x + p1y, c1y - p1x, pic1hDc)
            c3& = pPoint(c1x - p1y, c1y + p1x, pic1hDc)
            If c0& <> -1 Then SetPixel pic2hDc, c2x + p2x, c2y + p2y, c0&
            If c1& <> -1 Then SetPixel pic2hDc, c2x - p2x, c2y - p2y, c1&
            If c2& <> -1 Then SetPixel pic2hDc, c2x + p2y, c2y - p2x, c2&
            If c3& <> -1 Then SetPixel pic2hDc, c2x - p2y, c2y + p2x, c3&
        Next

        
        
    Next

End Sub

Sub bmp_rotate3(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Double)

    'Rotate the image in a picture box.

    'pic1 is the picture box with the bitmap to rotate

    'pic2 is the picture box to receive the rotated bitmap

    'theta is the angle of rotation

    
    Dim c1x As Double, c1y As Double
    Dim c2x As Double, c2y As Double
    Dim a As Double
    Dim p1x As Double, p1y As Double
    Dim p2x As Double, p2y As Double
    Dim n As Integer, r As Double
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = (n - 1) * 2
    pic1hDc = pic1.hdc
    pic2hDc = pic2.hdc


    For p2x = 0 To n / 2


        For p2y = 0 To n / 2
            If p2x = 0 Then a = Pi / 2 Else a = Sin(p2y / p2x) 'Try some things like Cos, Tan, Sin, 1/Tan, 1/Sin etc., etc.
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0& = pPoint(c1x + p1x, c1y + p1y, pic1hDc)
            c1& = pPoint(c1x - p1x, c1y - p1y, pic1hDc)
            c2& = pPoint(c1x + p1y, c1y - p1x, pic1hDc)
            c3& = pPoint(c1x - p1y, c1y + p1x, pic1hDc)
            If c0& <> -1 Then SetPixel pic2hDc, c2x + p2x, c2y + p2y, c0&
            If c1& <> -1 Then SetPixel pic2hDc, c2x - p2x, c2y - p2y, c1&
            If c2& <> -1 Then SetPixel pic2hDc, c2x + p2y, c2y - p2x, c2&
            If c3& <> -1 Then SetPixel pic2hDc, c2x - p2y, c2y + p2x, c3&
        Next

        
        
    Next

End Sub

'Use this function to interpolate between up to 4 pixels
Function pPoint(ByVal x As Double, ByVal y As Double, ByVal obj As Long) As Long
    z = Int(x)
    z2 = Int(x + 0.999)
    d2 = x - z
    If (z - z2) = 0 Then 'X integer
        pPoint = pPoint2(z, y, obj)
    Else
        pPoint = RGBDiv(pPoint2(z, y, obj), d2, pPoint2(z2, y, obj), (1 - d2))
    End If
End Function

Function pPoint2(ByVal x As Double, ByVal y As Double, ByVal obj As Long) As Long
    z = Int(y)
    z2 = Int(y + 0.999)
    d2 = y - z
    If (z - z2) = 0 Then 'Y integer
        pPoint2 = GetPixel(obj, x, z)
    Else
        pPoint2 = RGBDiv(GetPixel(obj, x, z), d2, GetPixel(obj, x, z2), (1 - d2))
    End If
End Function

Function RGBDiv(ByVal c1 As Long, ByVal p2 As Double, ByVal c2 As Long, ByVal p1 As Double)
    r1 = c1 And 255
    g1 = (c1 And (256 ^ 2 - 256)) / 256
    b1 = (c1 And (256 ^ 3 - 65536)) / (256 ^ 2)
    r2 = c2 And 255
    'If g1 = 255 And b1 = 0 Then Stop
    g2 = (c2 And (256 ^ 2 - 256)) / 256
    b2 = (c2 And (256 ^ 3 - 65536)) / (256 ^ 2)
    R3 = r1 * p1 + r2 * p2
    G3 = g1 * p1 + g2 * p2
    B3 = b1 * p1 + b2 * p2
    RGBDiv = RGB(R3, G3, B3)
End Function
