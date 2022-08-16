Public Class ExcelStyle
    Public Property Hue As Single = 100.0F
    Public Property Saturation As Single = 138.0F
    Public Property Luminesence As Single = 162.0F
    Public Property Font As Int16 = 0
    Public Property Shade As Int16

    Public Property BorderColor As String = "#000000"
    Public Property BorderWeight As String = "1"

    Public Property Color As String
        Get
            Dim c = ColorConverter.HslToRgba(Hue / 255.0F, Saturation / 255.0F, Luminesence / 255.0F)
            Return String.Format("{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B)
        End Get
        Set(value As String)
            Hue = Convert.ToInt16(value.Substring(0, 2), 16)
            Saturation = Convert.ToInt16(value.Substring(2, 2), 16) / 255.0F
            Luminesence = Convert.ToInt16(value.Substring(4, 2), 16) / 255.0F
        End Set
    End Property

    Public Function ToString() As String
        Return "s_h" & Hue & "_f" & Font & "_s" & Shade
    End Function
End Class
