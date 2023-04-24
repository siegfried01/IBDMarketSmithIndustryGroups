Public Class BadCSVFileException
    Inherits System.Exception
    Public Sub New(fn As String)
        MyBase.New()
        Me.FileName = fn
    End Sub

    Public FileName As String
End Class
