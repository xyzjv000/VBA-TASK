Sub ExportChartToHTML()
    Dim summaryWs As Worksheet
    Dim combinedWs As Worksheet
    Dim marginOnlyWs As Worksheet
    Dim actualCost As ChartObject
    Dim htmlTemplate As String
    Dim htmlContent As String
    Dim filePath As String
    Dim summaryData As Range
    Dim combinedData As Range
    Dim marginData As Range
    Dim summaryJson As String
    Dim jsonData As String
    Dim templateFilePath As String
    Dim templateFileNumber As Integer
    Dim jsonResult As String
    jsonResult = ExtractFilteredDataToJSONArray("FINAL output 2", "PivotTable1")
    Debug.Print jsonResult
    ' Set the worksheet
    Set summaryWs = ThisWorkbook.Sheets("Summary")
    Set combinedWs = ThisWorkbook.Sheets("Combined")
    Set marginOnlyWs = ThisWorkbook.Sheets("Retail Margin Only")
    ' Define the path to the HTML template file
    templateFilePath = GetCurrentExcelDirectory & "\Exports\HTML_Template.html" ' Change to your template file path
    
    ' Read HTML template from file
    templateFileNumber = FreeFile
    Open templateFilePath For Input As #templateFileNumber
    htmlTemplate = Input$(LOF(templateFileNumber), templateFileNumber)
    Close #templateFileNumber
    
    ' Define the range of data to be exported
    ' Set combinedData = Union(combinedWs.Range("C4:C" & combinedWs.Cells(Rows.Count, "C").End(xlUp).Row), _
    '                          combinedWs.Range("D4:D" & combinedWs.Cells(Rows.Count, "D").End(xlUp).Row), _
    '                          combinedWs.Range("E4:E" & combinedWs.Cells(Rows.Count, "E").End(xlUp).Row), _
    '                          combinedWs.Range("F4:F" & combinedWs.Cells(Rows.Count, "F").End(xlUp).Row), _
    '                          combinedWs.Range("G4:G" & combinedWs.Cells(Rows.Count, "G").End(xlUp).Row), _
    '                          combinedWs.Range("H4:H" & combinedWs.Cells(Rows.Count, "H").End(xlUp).Row))
    Set summaryData = summaryWs.Range("A7").CurrentRegion
    
    ' Convert the range to JSON
    summaryJson = jsonResult
    ' combinedJson = RangeToJSON(combinedData)
    
    ' Replace placeholders in the HTML template with actual data
    
    ' htmlContent = Replace(htmlTemplate, "{{combinedJson}}", combinedJson)
    ' htmlContent = Replace(htmlContent, "{{summaryJson}}", summaryJson)
    htmlContent = Replace(htmlTemplate, "{{summaryJson}}", summaryJson)
    ' Define the file path to save the HTML file
    filePath = GetCurrentDesktopirectory & "\ExportedReport.html" ' Change to your desired file path
    
    ' Write HTML content to file
    Open filePath For Output As #1
    Print #1, htmlContent
    Close #1
    
    ' Notify the user
    MsgBox "Data and chart have been exported to HTML successfully!", vbInformation
End Sub

Function GetCurrentExcelDirectory() As String
    ' Get the directory of the currently open workbook
    Dim currentDirectory As String
    currentDirectory = ThisWorkbook.Path
    
    ' Check if the workbook is saved
    If currentDirectory = "" Then
        GetCurrentExcelDirectory = "The workbook has not been saved yet."
    Else
        GetCurrentExcelDirectory = currentDirectory
    End If
End Function

Function GetCurrentDesktopirectory() As String
    ' Get the directory of the currently open workbook
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    ' Check if the workbook is saved
    If desktopPath = "" Then
        GetCurrentDesktopirectory = "The workbook has not been saved yet."
    Else
        GetCurrentDesktopirectory = desktopPath
    End If
End Function


Function RangeToJSON(rng As Range) As String
    Dim data As String
    Dim i As Integer, j As Integer
    Dim headers As Variant
    Dim values As Variant
    Dim value As Variant
    Dim headerName As String

    On Error GoTo ErrorHandler

    headers = rng.Rows(1).Value
    values = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).Value

    data = "["

    For i = 1 To UBound(values, 1)
        data = data & "{"
        For j = 1 To UBound(values, 2)
            value = Trim(CStr(values(i, j))) ' Ensure value is a string and remove spaces
            headerName = ToCamelCase(CStr(headers(1, j)))
            data = data & """" & headerName & """: """ & value & """"
            If j < UBound(values, 2) Then
                data = data & ", "
            End If
        Next j
        data = data & "}"
        If i < UBound(values, 1) Then
            data = data & ", "
        End If
    Next i

    data = data & "]"
    
    RangeToJSON = data
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    RangeToJSON = "[]"
End Function

Function ToCamelCase(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim upperNext As Boolean

    result = ""
    upperNext = False

    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If char Like "[A-Z]" Then
            If i = 1 Then
                result = result & LCase(char)
            Else
                result = result & IIf(upperNext, LCase(char), char)
                upperNext = False
            End If
        ElseIf char Like "[a-z]" Then
            result = result & char
            upperNext = False
        ElseIf char Like " " Then
            upperNext = True
        End If
    Next i

    ToCamelCase = result
End Function

' Sub ExtractFilteredDataToJSONArray()
'     Dim pt As PivotTable
'     Dim ws As Worksheet
'     Dim jsonString As String
'     Dim jsonArray As Object
'     Dim pRow As Range
'     Dim keyName As String
'     Dim value As String
'     Dim data As Object

'     ' Set worksheet and PivotTable
'     Set ws = ThisWorkbook.Worksheets("FINAL output 2") ' Adjust the sheet name
'     Set pt = ws.PivotTables("PivotTable1") ' Adjust the PivotTable name
    
'     ' Create JSON array
'     Set jsonArray = CreateObject("Scripting.Dictionary")
'     Set data = CreateObject("Scripting.Dictionary")
    
'     ' Get visible row data
'     For Each pRow In pt.DataBodyRange.Rows
'         keyName = LCase(pRow.Cells(1, 0).Value)  ' The value of the first column will be the key
'         value = pRow.Cells(1, 1).Value    ' The value of the second column will be the value
'         data.Add keyName, value           ' Add to dictionary
'     Next pRow
'     ' Convert to JSON string
'     jsonString = JsonConvertToArray(data)
    
'     ' Output JSON string to Immediate window
'     Debug.Print jsonString
' End Sub

' Function JsonConvertToArray(dict As Object) As String
'     Dim json As String
'     Dim key As Variant
'     Dim i As Integer
    
'     json = "["
'     i = 0
'     For Each key In dict
'         json = json & "{" & """" & key & """:""" & dict(key) & """}"
'         If i < dict.Count - 1 Then json = json & ","
'         i = i + 1
'     Next key
'     json = json & "]"
    
'     JsonConvertToArray = json
' End Function

Function ExtractFilteredDataToJSONArray(sheetName As String, pivotTableName As String) As String
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim jsonString As String
    Dim jsonArray As Object
    Dim data As Object
    Dim typesofmargin As String
    Dim totalName As String
    Dim totalValue As String
    Dim rowDict As Object
    Dim totalDict As Object
    Dim pRow As Range
    Dim pCol As Range
    Dim colHeaders As Range
    Dim startCol As Integer
    Dim i As Integer

    ' Set worksheet and PivotTable
    Set ws = ThisWorkbook.Worksheets(sheetName) ' Adjust the sheet name
    Set pt = ws.PivotTables(pivotTableName) ' Adjust the PivotTable name

    ' Create JSON array
    Set jsonArray = CreateObject("Scripting.Dictionary")

    ' Get column headers range
    Set colHeaders = pt.DataBodyRange.Cells(1, 1).Offset(-1, 1).Resize(1, pt.DataBodyRange.Columns.Count - 1)
    startCol = pt.DataBodyRange.Column ' First data column

    ' Iterate over the PivotTable columns, starting from the second column
    For i = 1 To pt.DataBodyRange.Columns.Count
        ' Initialize a new row dictionary for each typesofmargin
        Set rowDict = CreateObject("Scripting.Dictionary")
        Set totalDict = CreateObject("Scripting.Dictionary")

        ' Get the typesofmargin from the column header
        typesofmargin = Trim(colHeaders.Cells(1, i - 1).Value)
        rowDict.Add "typesofmargin", typesofmargin

        ' Loop through the rows to get totals for this typesofmargin
        For Each pRow In pt.DataBodyRange.Rows
            totalName = LCase(Trim(pRow.Cells(1, 0).Value)) ' The first column under "Row Labels"
            totalValue = pRow.Cells(1, i).Value ' The current column value for this row            
            If totalName = "grand total" Then
                totalName = "total"                
            End If
            totalDict.Add totalName, totalValue
        Next pRow

        ' Add totals dictionary to row dictionary
        rowDict.Add "totals", totalDict
        ' Add row dictionary to JSON array
        jsonArray.Add jsonArray.Count, rowDict
    Next i

    ' Convert to JSON string
    jsonString = JsonConvertToObject(jsonArray)
    ExtractFilteredDataToJSONArray = jsonString
End Function

Function JsonConvertToObject(dict As Object) As String
    Dim json As String
    Dim key As Variant
    Dim subKey As Variant
    Dim i As Integer
    Dim subJson As String

    json = "["
    i = 0
    For Each key In dict
        json = json & "{"
        json = json & """typesofmargin"":""" & dict(key).Item("typesofmargin") & ""","

        ' Convert totals dictionary to JSON object
        subJson = "{"
        For Each subKey In dict(key).Item("totals")
            subJson = subJson & """" & subKey & """:""" & dict(key).Item("totals")(subKey) & ""","
        Next subKey
        ' Remove the trailing comma
        If Right(subJson, 1) = "," Then subJson = Left(subJson, Len(subJson) - 1)
        subJson = subJson & "}"

        json = json & """totals"":" & subJson & "}"
        If i < dict.Count - 1 Then json = json & ","
        i = i + 1
    Next key
    json = json & "]"

    JsonConvertToObject = json
End Function
