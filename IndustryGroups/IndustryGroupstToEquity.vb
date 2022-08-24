﻿Option Strict Off
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports CsvHelper
Imports CsvHelper.Configuration
Imports Microsoft.VisualBasic.FileIO

Public Class MissingFile
    Inherits System.Exception
    Public Property FileName As String
    Public Sub New(fileName As String)
        MyBase.New("File not found: " & fileName)
        Me.FileName = fileName
    End Sub
End Class
Public Class IndustryGroupstToEquity

    ' static function to cvs file
    Public Shared Function LoadTable(fileName As String, Optional maxDaysOld As Int16 = 4) As Dictionary(Of String, List(Of Equity))
        Dim displayFileName As String = fileName
        fileName = Environment.ExpandEnvironmentVariables(fileName)
        IsFileTooOld(fileName, maxDaysOld, displayFileName)
        Dim result = New Dictionary(Of String, List(Of Equity))
        Dim industryName As String
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
            industryName = fields(colNames("Industry Name"))
            symbol = fields(colNames("Symbol"))
            compRating = ParseField(colNames, fields, "Comp Rating")
            Dim price = Double.Parse(fields(colNames("Current Price")), CultureInfo.InvariantCulture)
            Dim dvol = Double.Parse(fields(colNames("50-Day Avg $ Vol (1000s)")), CultureInfo.InvariantCulture)
            Dim rs = ParseField(colNames, fields, "RS Rating") 'Double.Parse(fields(colNames("RS Rating")), CultureInfo.InvariantCulture)
            Dim smr = fields(colNames("SMR Rating"))
            Dim ad = fields(colNames("A/D Rating"))
            Dim yield = Double.Parse(fields(colNames("Yield %")), CultureInfo.InvariantCulture)
            Dim eps = ParseField(colNames, fields, "EPS Rating") 'Double.Parse(fields(colNames("EPS Rating")), CultureInfo.InvariantCulture)
            Dim upDown = ParseField(colNames, fields, "Up/Down Vol")
            Dim name = fields(colNames("Name"))
            Dim eq = New Equity()
            eq.TickerSymbol = symbol
            eq.Composite = compRating
            eq.Price = price
            eq.DollarVolume = dvol
            eq.RS = rs
            eq.SMR = smr
            eq.Yield = yield
            eq.EPS = eps
            eq.AD = ad
            eq.UpDown = upDown
            eq.Name = name
            If result.ContainsKey(industryName) Then
                result(industryName).Add(eq)
            Else
                result.Add(industryName, New List(Of Equity) From {eq})
            End If
        End While
        Return result
    End Function

    Private Shared Function ParseField(colNames As Dictionary(Of String, Integer), fields() As String, fieldName As String) As Double
        Dim ratingValue As Double
        Dim ratingStr = fields(colNames(fieldName))
        If String.IsNullOrEmpty(ratingStr) Or ratingStr = "-" Then
            ratingStr = "0"
        End If
        ratingValue = Double.Parse(ratingStr, CultureInfo.InvariantCulture)
        Return ratingValue
    End Function
End Class
