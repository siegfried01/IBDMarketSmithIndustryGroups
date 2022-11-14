Public Class ExcelStyle
    Private hueValue = 100.0F
    Public Property Hue As Single
        Get
            Return hueValue * HueScale + HueOffset ' use linear mapping to compute hue based some metric like composite value. HueOffest & HueScale were perviously computed from the composit rating
        End Get
        Set(value As Single)
            hueValue = value
        End Set
    End Property
    Public Property HueScale As Single = 2.3F
    Public Property HueOffset As Single = 93.3F
    Public Property Saturation As Single = 200.0F
    Public Property Luminesence As Single = 200.0F
    Public Property Font As Int16 = 0
    Public Property Shade As Int16
    Private ymax As Single
    Public Property HueMax As Single
        Get
            Return ymax
        End Get
        Set(value As Single)
            ymax = value
            recalc()
        End Set
    End Property
    Private ymin As Single
    Public Property HueMin As Single
        Get
            Return ymin
        End Get
        Set(value As Single)
            ymin = value
            recalc()
        End Set
    End Property
    Private xmax As Single
    Public Property InputMetricMax As Single
        Get
            Return xmax
        End Get
        Set(value As Single)
            xmax = value
            recalc()
        End Set
    End Property

    Private Sub recalc()
        HueScale = (HueMin + (xmax * HueMin / xmin - HueMax) / (1 - xmax / xmin)) / xmin
        HueOffset = (HueMax - xmax * HueMin / xmin) / (1 - xmax / xmin)
    End Sub

    Private xmin As Single
    Public Property InputMetricMin As Single
        Get
            Return xmin
        End Get
        Set(value As Single)
            xmin = value
            recalc()
        End Set
    End Property
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

    Public Overrides Function ToString() As String
        Return "s_h" & Hue & "_f" & Font & "_s" & Shade
    End Function
End Class
