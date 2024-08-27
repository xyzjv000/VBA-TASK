Sub ExportChartToHTML()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim htmlTemplate As String
    Dim htmlContent As String
    Dim filePath As String
    Dim chartImagePath As String
    Dim dataRange As Range
    Dim cell As Range
    Dim data As String
    Dim i As Integer, j As Integer
    Dim templateFilePath As String
    Dim templateFileNumber As Integer
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Summary") ' Change "Sheet1" to your sheet name
    
    ' Save chart as an image
    Set chartObj = ws.ChartObjects("Chart 1") ' Change "Chart 1" to your chart name
    chartImagePath = GetCurrentExcelDirectory & "\Exports\Images\chart_image.png" ' Change to your desired path
    chartObj.Chart.Export Filename:=chartImagePath, FilterName:="PNG"
    
    ' Define the path to the HTML template file
    templateFilePath = GetCurrentExcelDirectory & "\Exports\HTML_Template.html" ' Change to your template file path
    
    ' Read HTML template from file
    templateFileNumber = FreeFile
    Open templateFilePath For Input As #templateFileNumber
    htmlTemplate = Input$(LOF(templateFileNumber), templateFileNumber)
    Close #templateFileNumber
    
    ' Initialize data string
    data = ""
    
    ' Define the range of data to be exported
    Set dataRange = ws.Range("A5").CurrentRegion ' Adjust the range as needed
    
    ' Loop through the range and build HTML table rows
    For i = 1 To dataRange.Rows.Count
        data = data & "<tr>"
        For j = 1 To dataRange.Columns.Count
            data = data & "<td>" & dataRange.Cells(i, j).Value & "</td>"
        Next j
        data = data & "</tr>"
    Next i
    
    ' Replace placeholders in the HTML template with actual data
    htmlContent = Replace(htmlTemplate, "{{chartImage}}", chartImagePath)
    htmlContent = Replace(htmlContent, "{{tableData}}", data)
    
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