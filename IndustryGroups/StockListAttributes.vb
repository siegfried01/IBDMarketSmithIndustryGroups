Public Class StockListAttributes
    Public Property Annotation As String = ""
    Public Property ExcelColumn As Integer = -1
    Public Property ExcelStyle As String = ""
    Public Property DisplayName() As String = ""
    Public Property CsvFileFoundAndLoaded() As Boolean = False
End Class
Public Class StockList
    Public Property Name As String = ""
    Public Property Attributes As StockListAttributes = New StockListAttributes
End Class
