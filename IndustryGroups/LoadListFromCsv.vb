Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.FileIO
Imports System.Xml
Imports System.Xml.XPath
Imports System.Console

Module LoadListFromCsv
    Function LoadIndustryGroups(ByRef industryGroups As XDocument, ByRef groupRows As IEnumerable(Of XElement), ss As XNamespace, hrefStyle As String, Optional maxDaysOld As Int16 = 4, Optional fileName As String = "%USERPROFILE%\Downloads\197 Industry Groups.csv") As IEnumerable(Of XElement)
        DeleteIndustryGroupDataRows(groupRows)
        Dim displayFileName = fileName
        fileName = Environment.ExpandEnvironmentVariables(fileName)
        IsFileTooOld(fileName, maxDaysOld, displayFileName)
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
        Dim rowCount = 1
        While tfp.EndOfData = False
            Dim fields = tfp.ReadFields()
            symbol = fields(colNames("Symbol"))
            Dim row = New XElement(ss + "Row", New XAttribute(ss + "AutoFitHeight", "0"))
            ' @@todo@@ update the excel column headers on first time thru so we can add new columns in a single place.
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Order"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), fields(colNames("Symbol"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "String"), fields(colNames("Name"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Number of Stocks"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Ind Group Rank"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Ind Grp Rnk Last Week"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Ind Grp Rnk 3 Mo Ago"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Ind Grp Rnk 6 Mo Ago"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("% Chg YTD"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("Ind Mkt Val (bil)"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("% Chg Cur Week"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))
            row.Add(New XElement(ss + "Cell", New XElement(ss + "Data", New XAttribute(ss + "Type", "Number"), fields(colNames("% Chg 1 Month"))), New XElement(ss + "NamedCell", New XAttribute(ss + "Name", "_FilterDatabase"))))

            Dim b = industryGroups.Element(ss + "Workbook")
            Dim w = b.Element(ss + "Worksheet")
            Dim t = w.Element(ss + "Table")
            t.Add(row)
        End While
        Return groupRows
    End Function

    Private Sub DeleteIndustryGroupDataRows(groupRows As IEnumerable(Of XElement))
        For rowCount = groupRows.Count - 1 To 1 Step -1
            groupRows(rowCount).Remove()
        Next
    End Sub
    Sub IsFileTooOld(ByVal fileName As String, ByVal days As Integer, ByVal displayFileName As String)
        Dim fileInfo As New FileInfo(fileName)
        If fileInfo.Exists Then
            Dim fileDate = fileInfo.LastWriteTime
            Dim now = DateTime.Now
            Dim diff = now.Subtract(fileDate)
            If diff.TotalDays > days Then
                ' delete the file
                'File.Delete(fileName)
                Throw New Exception("File is too old: " & displayFileName & " File has been deleted.")
            End If
        Else
            Throw New MissingFile(fileName)
        End If
    End Sub

    Function LoadListFromCsv(Optional fileName As String = "%USERPROFILE%\Downloads\IBD Live Ready.csv", Optional maxDaysOld As Int16 = 4) As HashSet(Of String)
        Dim result As New HashSet(Of String)
        Dim displayFileName = fileName
        fileName = Environment.ExpandEnvironmentVariables(fileName)

        IsFileTooOld(fileName, maxDaysOld, displayFileName)
        Dim symbol As String
        Dim tfp As New TextFieldParser(fileName)
        tfp.Delimiters = New String() {","}
        tfp.TextFieldType = FieldType.Delimited

        'Try
        Dim colArray = tfp.ReadLine().Split(","c).ToList().Select(Of String)(Function(x) x.Substring(1, x.Length - 2).ToArray())
        Dim colNames = New Dictionary(Of String, Integer)
        Dim i = 0
        Dim j = 0
        For Each col In colArray
            If colNames.Keys.Contains(col) Then
                j = j + 1
                colNames.Add(col & "_" & j, i)
            Else
                colNames.Add(col, i)
            End If
            i = i + 1
        Next
        While tfp.EndOfData = False
            Dim fields = tfp.ReadFields()
            Try
                symbol = fields(colNames("Symbol")) ' @@todo@@ throw an exception here if this fails because it means we have a bogus file that can be ignored.
            Catch ex As System.Collections.Generic.KeyNotFoundException
                WriteLine("Key Not found exception: file=" & fileName)
                Throw New BadCSVFileException(fileName)
            End Try
            result.Add(symbol)
        End While
        'Catch ex As Exception
        'Throw New Exception("Error loading file: " & displayFileName & " " & ex.Message)
        'End Try
        Return result
    End Function
End Module
