Sub ExportChartToHTML()
    Dim ws As Worksheet
    Dim actualCost As ChartObject
    Dim htmlTemplate As String
    Dim htmlContent As String
    Dim filePath As String
    Dim actualCostPath As String
    Dim dataRange As Range
    Dim jsonData As String
    Dim templateFilePath As String
    Dim templateFileNumber As Integer
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Summary") ' Change "Summary" to your sheet name
    
    ' Save chart as an image
    Set actualCost = ws.ChartObjects("Actual Cost")

    actualCostPath = GetCurrentExcelDirectory & "\Exports\Images\actualCost.png" ' Change to your desired path
    actualCost.Chart.Export Filename:=actualCostPath, FilterName:="PNG"
    
    ' Define the path to the HTML template file
    templateFilePath = GetCurrentExcelDirectory & "\Exports\HTML_Template.html" ' Change to your template file path
    
    ' Read HTML template from file
    templateFileNumber = FreeFile
    Open templateFilePath For Input As #templateFileNumber
    htmlTemplate = Input$(LOF(templateFileNumber), templateFileNumber)
    Close #templateFileNumber
    
    ' Define the range of data to be exported
    Set dataRange = ws.Range("A7").CurrentRegion ' Adjust the range as needed
    
    ' Convert the range to JSON
    jsonData = RangeToJSON(dataRange)
    
    ' Replace placeholders in the HTML template with actual data
    htmlContent = Replace(htmlTemplate, "{{actualCostPath}}", actualCostPath)
    htmlContent = Replace(htmlContent, "{{jsonData}}", jsonData)
    
    ' Define the file path to save the HTML file
    filePath = GetCurrentExcelDirectory & "\Exports\ExportedData.html" ' Change to your desired file path
    
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
