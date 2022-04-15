Option Strict Off
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports CsvHelper
Imports CsvHelper.Configuration
Imports Microsoft.VisualBasic.FileIO

Public Class IndustryGroupstToEquity
    ' static function to cvs file
    Public Shared Function LoadTable(Optional fileName As String = "%DN%\MinDollarVol20MComp80.csv") As Dictionary(Of String, List(Of (TickerSymbol As String, comp As Double)))
        fileName = Environment.ExpandEnvironmentVariables(fileName)
        Dim result = New Dictionary(Of String, List(Of (TickerSymbol As String, comp As Double)))
        Dim name As String
        Dim symbol As String
        Dim compRating As Double

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
            name = fields(colNames("Industry Name"))
            symbol = fields(colNames("Symbol"))
            compRating = Double.Parse(fields(colNames("Comp Rating")), CultureInfo.InvariantCulture)
            If result.ContainsKey(name) Then
                result(name).Add((symbol, compRating))
            Else
                result.Add(name, New List(Of (TickerSymbol As String, comp As Double)) From {(symbol, compRating)})
            End If
        End While
        Return result
    End Function


End Class
