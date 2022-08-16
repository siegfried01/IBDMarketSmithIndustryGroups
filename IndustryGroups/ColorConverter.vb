Imports System.Drawing
' https://stackoverflow.com/questions/2353211/hsl-to-rgb-color-conversion
Public Class ColorConverter
    Public Shared Function HslToRgba(hue As Single, sat As Single, lum As Single) As Color
        Dim r As Single, g As Single, b As Single
        If sat = 0 Then
            r = lum
            g = lum
            b = lum
        Else
            Dim q As Single = If(lum < 0.5, lum * (1 + sat), lum + sat - lum * sat)
            Dim p As Single = 2 * lum - q
            r = HslToRgb(p, q, hue + 1 / 3)
            g = HslToRgb(p, q, hue)
            b = HslToRgb(p, q, hue - 1 / 3)
        End If
        Return Color.FromArgb(255, CInt(r * 255), CInt(g * 255), CInt(b * 255))
    End Function
    Public Shared Function ToHexString(c As Color) As String
        Return String.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B)
    End Function

    Private Shared Function HslToRgb(p As Single, q As Single, t As Single) As Single
        If t < 0 Then t += 1
        If t > 1 Then t -= 1
        If t < 1 / 6 Then Return p + (q - p) * 6 * t
        If t < 1 / 2 Then Return q
        If t < 2 / 3 Then Return p + (q - p) * (2 / 3 - t) * 6
        Return p
    End Function

End Class
