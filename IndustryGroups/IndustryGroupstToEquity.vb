Option Strict Off
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports CsvHelper
Imports CsvHelper.Configuration
Imports Microsoft.VisualBasic.FileIO

Public Class IndustryGroupstToEquity

    ' static function to cvs file
    Public Shared Function LoadTable(Optional fileName As String = "%USERPROFILE%\Downloads\MinDollarVol20MComp80.csv") As Dictionary(Of String, List(Of Equity))
        Dim displayFileName As String = fileName
        fileName = Environment.ExpandEnvironmentVariables(fileName)
        IsFileTooOld(fileName, 4, displayFileName)
        Dim result = New Dictionary(Of String, List(Of Equity))
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
            Dim price = Double.Parse(fields(colNames("Current Price")), CultureInfo.InvariantCulture)
            Dim dvol = Double.Parse(fields(colNames("50-Day Avg $ Vol (1000s)")), CultureInfo.InvariantCulture)
            Dim rs = Double.Parse(fields(colNames("RS Rating")), CultureInfo.InvariantCulture)
            Dim smr = fields(colNames("SMR Rating"))
            Dim ad = fields(colNames("A/D Rating"))
            Dim yield = Double.Parse(fields(colNames("Yield %")), CultureInfo.InvariantCulture)
            Dim eps = Double.Parse(fields(colNames("EPS Rating")), CultureInfo.InvariantCulture)
            Dim eq = New Equity()
            eq.TickerSymbol = symbol
            eq.Composite = compRating
            eq.Price = price
            eq.DollarVolume = dvol
            eq.RS = rs
            eq.SMR = smr
            eq.Yield = yield
            eq.EPS = eps
            If result.ContainsKey(name) Then
                result(name).Add(eq)
            Else
                result.Add(name, New List(Of Equity) From {eq})
            End If
        End While
        Return result
    End Function


End Class
