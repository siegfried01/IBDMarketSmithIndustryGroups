Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.FileIO

Module LoadListFromCsv
    Function LoadListFromCsv(Optional fileName As String = "%DN%\IBD Live Ready.csv") As HashSet(Of String)
        Dim result As New HashSet(Of String)
        fileName = Environment.ExpandEnvironmentVariables(fileName)
        Dim symbol As String
        Dim tfp As New TextFieldParser(fileName)
        tfp.Delimiters = New String() {","}
        tfp.TextFieldType = FieldType.Delimited


        Dim colArray = tfp.ReadLine().Split(","c).ToList().Select(Of String)(Function(x) x.Substring(1, x.Length - 2).ToArray())
        Dim colNames = New Dictionary(Of String, Integer)
        Dim i = 0
        For Each col In colArray
            colNames.Add(col, i)
            i = i + 1
        Next
        While tfp.EndOfData = False
            Dim fields = tfp.ReadFields()
            symbol = fields(colNames("Symbol"))
            result.Add(symbol)
        End While
        Return result
    End Function
End Module
