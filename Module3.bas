Sub ExportChartToHTML()
    Dim summaryWs As Worksheet
    Dim combinedWs As Worksheet
    Dim marginOnlyWs As Worksheet
    Dim actualCost As ChartObject
    Dim htmlTemplate As String
    Dim htmlContent As String
    Dim htmlWSummary As String
    Dim htmlWVersion As String
    Dim filePath As String
    Dim summaryData As Range
    Dim combinedData As Range
    Dim marginData As Range
    Dim combinedJson As String
    Dim summaryJson As String
    Dim nmiJson As String
    Dim combinedFullJson As String
    Dim jsonData As String
    Dim templateFilePath As String
    Dim templateFileNumber As Integer
    Dim jsonResult1 As String
    Dim jsonResult2 As String
    Dim jsonResult3 As String
    Dim jsonResult4 As String
    Dim versionData As String

    jsonResult1 = ExtractFilteredDataToJSONArrayMargins("FINAL output 2", "PivotTable1")
    jsonResult2 = ExtractFilteredDataToJSONArrayMargins("FINAL output 2", "PivotTable2")
    jsonResult3 = ExtractFilteredDataToJSONArrayMargins("Retail Margin Only", "PivotTable21")
    jsonResult4 = ExportPivotTableToJSON("FINAL output 2", "PivotTable4")
    ' templateFilePath = GetCurrentExcelDirectory & "\Exports\HTML_Template.html"
    templateFilePath = GetCurrentExcelDirectory & Application.PathSeparator & "Exports" & Application.PathSeparator & "HTML_Template.html"
    ' Read HTML template from file
    templateFileNumber = FreeFile
    Open templateFilePath For Input As #templateFileNumber
    htmlTemplate = Input$(LOF(templateFileNumber), templateFileNumber)
    Close #templateFileNumber

    ' Convert the range To JSON
    combinedJson = jsonResult1
    summaryJson = jsonResult2
    nmiJson = jsonResult3
    combinedFullJson = jsonResult4
    versionData = GetVersionData()
    Debug.Print combinedJson
    ' Replace placeholders in the HTML template With actual data
    htmlContent = Replace(htmlTemplate, "{{combinedJson}}", combinedJson)
    htmlContent = Replace(htmlContent, "{{summaryJson}}", summaryJson)
    htmlContent = Replace(htmlContent, "{{versionData}}", versionData)
    htmlContent = Replace(htmlContent, "{{nmiJson}}", nmiJson)
    htmlContent = Replace(htmlContent, "{{combinedFullJson}}", combinedFullJson)
    ' Define the file path To save the HTML file
    filePath = GetCurrentDesktopirectory & "\ExportedReport.html" ' Change To your desired file path

    ' Write HTML content To file
    Open filePath For Output As #1
    Print #1, htmlContent
    Close #1

    ' Notify the user
    MsgBox "Data And chart have been exported To HTML successfully!", vbInformation
End Sub

Function GetCurrentExcelDirectory() As String
    ' Get the directory of the currently open workbook
    Dim currentDirectory As String
    currentDirectory = ThisWorkbook.Path

    ' Check If the workbook is saved
    If currentDirectory = "" Then
        GetCurrentExcelDirectory = "The workbook has Not been saved yet."
    Else
        GetCurrentExcelDirectory = currentDirectory
    End If
End Function

Function GetCurrentDesktopirectory() As String
    ' Get the directory of the currently open workbook
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    ' Check If the workbook is saved
    If desktopPath = "" Then
        GetCurrentDesktopirectory = "The workbook has Not been saved yet."
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

    On Error Goto ErrorHandler

        headers = rng.Rows(1).value
        values = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).value

        data = "["

        For i = 1 To UBound(values, 1)
            data = data & "{"
            For j = 1 To UBound(values, 2)
                value = Trim(CStr(values(i, j))) ' Ensure value is a string And remove spaces
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
        Elseif char Like "[a-z]" Then
            result = result & char
            upperNext = False
        Elseif char Like " " Then
            upperNext = True
        End If
    Next i

    ToCamelCase = result
End Function

Function ExtractFilteredDataToJSONArrayMargins(sheetName As String, pivotTableName As String) As String
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim jsonString As String
    Dim jsonArray As Object
    Dim rowDict As Object
    Dim totalDict As Object
    Dim labelDict As Object
    Dim pRow As Range
    Dim colHeaders As Range
    Dim startCol As Integer
    Dim typesofmargin As String
    Dim totalName As String
    Dim totalLabel As String
    Dim totalValue As String
    Dim i As Integer

    ' Set worksheet And PivotTable
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set pt = ws.PivotTables(pivotTableName)

    ' Create JSON array
    Set jsonArray = CreateObject("Scripting.Dictionary")

    ' Get column headers range
    Set colHeaders = pt.DataBodyRange.Cells(1, 1).Offset(-1, 1).Resize(1, pt.DataBodyRange.Columns.Count - 1)
    startCol = pt.DataBodyRange.Column

    ' Iterate over the PivotTable columns, starting from the second column
    For i = 1 To pt.DataBodyRange.Columns.Count
        ' Initialize a New row dictionary For each typesofmargin
        Set rowDict = CreateObject("Scripting.Dictionary")
        Set totalDict = CreateObject("Scripting.Dictionary")
        Set labelDict = CreateObject("Scripting.Dictionary")

        ' Get the typesofmargin from the column header
        typesofmargin = Trim(colHeaders.Cells(1, i - 1).value)
        rowDict.Add "typesofmargin", typesofmargin

        ' Loop through the rows To Get totals For this typesofmargin
        For Each pRow In pt.DataBodyRange.Rows
            totalName = LCase(Trim(pRow.Cells(1, 0).value)) ' The first column under "Row Labels"
            totalLabel = pRow.Cells(1, 0).value ' Use the correct reference For labels
            totalValue = pRow.Cells(1, i).value ' The current column value For this row

            ' Normalize "grand total" To "total"
            If totalName = "grand total" Then
                totalLabel = "Total"
                totalName = "total"
            End If

            ' Add total And Label To respective dictionaries
            totalDict.Add totalName, totalValue
            labelDict.Add totalName, totalLabel
        Next pRow

        ' Add totals And labels dictionaries To row dictionary
        rowDict.Add "totals", totalDict
        rowDict.Add "labels", labelDict

        ' Add row dictionary To JSON array
        jsonArray.Add jsonArray.Count, rowDict
    Next i

    ' Convert To JSON string
    jsonString = JsonConvertToObjectMargins(jsonArray)
    ExtractFilteredDataToJSONArrayMargins = jsonString
End Function

Function JsonConvertToObjectMargins(dict As Object) As String
    Dim json As String
    Dim key As Variant
    Dim subKey As Variant
    Dim i As Integer
    Dim subJson As String
    Dim labelJson As String

    json = "["
    i = 0
    For Each key In dict
        json = json & "{"
        json = json & """typesofmargin"":""" & dict(key).Item("typesofmargin") & ""","

        ' Convert totals dictionary To JSON object
        subJson = "{"
        For Each subKey In dict(key).Item("totals")
            subJson = subJson & """" & subKey & """:""" & dict(key).Item("totals")(subKey) & ""","
        Next subKey
        ' Remove the trailing comma from totals
        If Right(subJson, 1) = "," Then subJson = Left(subJson, Len(subJson) - 1)
            subJson = subJson & "}"

            ' Convert labels dictionary To JSON object
            labelJson = "{"
            For Each subKey In dict(key).Item("labels")
                labelJson = labelJson & """" & subKey & """:""" & dict(key).Item("labels")(subKey) & ""","
            Next subKey
            ' Remove the trailing comma from labels
            If Right(labelJson, 1) = "," Then labelJson = Left(labelJson, Len(labelJson) - 1)
                labelJson = labelJson & "}"

                ' Combine totals And labels into the main JSON object
                json = json & """totals"":" & subJson & "," & """labels"":" & labelJson & "}"
                If i < dict.Count - 1 Then json = json & ","
                    i = i + 1
                Next key
                json = json & "]"

                JsonConvertToObjectMargins = json
End Function


Function GetVersionData() As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim jsonString As String
    Dim i As Integer

    On Error Goto ErrorHandler

        ' Set the worksheet And range
        Set ws = ThisWorkbook.Sheets("Run Sheet")
        Set rng = ws.Range("B20:C25")

        ' Start building the JSON string
        jsonString = "["

        ' Loop through each row in the range And build JSON
        For i = 1 To rng.Rows.Count
            If Not IsEmpty(rng.Cells(i, 1)) And Not IsEmpty(rng.Cells(i, 2)) Then
                jsonString = jsonString & "{"
                jsonString = jsonString & """version"": """ & rng.Cells(i, 1).value & """, "
                jsonString = jsonString & """effectiveDate"": """ & rng.Cells(i, 2).value & """"
                jsonString = jsonString & "}"

                ' Add a comma If this is Not the last item
                If i < rng.Rows.Count Then
                    jsonString = jsonString & ", "
                End If
            End If
        Next i

        ' Close the JSON array
        jsonString = jsonString & "]"

        ' Return the JSON string
        GetVersionData = jsonString
     Exit Function

 ErrorHandler:
        GetVersionData = "Error: " & Err.Description
End Function

Function ExportPivotTableToJSON(sheetName As String, pivotTableName As String) As String
    Dim ws As Worksheet
    Dim pvt As PivotTable
    Dim jsonData As String
    Dim rowItem As PivotItem
    Dim rowItemNmi As Variant
    Dim rowItemCapacity As Variant
    Dim rowItemCommission As Variant
    Dim rowItemEss As Variant
    Dim rowItemLgc As Variant
    Dim rowItemMarketFees As Variant
    Dim rowItemNetwork As Variant
    Dim rowItemRetailMargin As Variant
    Dim rowItemRevenue As Variant
    Dim rowItemStc As Variant
    Dim rowItemWholesaleEnergy As Variant
    Dim rowItemPortfolio As Variant
    Dim rowItemStatus As Variant
    Dim rowItemAssociation As Variant
    Dim rowItemAgreement As Variant
    Dim colItem As PivotItem
    Dim valueNmi As Variant
    Dim valueCapacity As Variant
    Dim valueCommission As Variant
    Dim valueEss As Variant
    Dim valueLgc As Variant
    Dim valueMarketFees As Variant
    Dim valueNetwork As Variant
    Dim valueRetailMargin As Variant
    Dim valueRevenue As Variant
    Dim valueStc As Variant
    Dim valueWholesaleEnergy As Variant
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim colFieldName As String
    Dim dataFieldName As String

    ' Set worksheet And pivot table
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set pvt = ws.PivotTables(pivotTableName)


    'Add 50 in Loop
    Dim nmiIndex As Long
    Dim portfolioIndex As Long
    Dim statusIndex As Long
    Dim associationIndex As Long
    Dim agreementIndex As Long
    Dim capacityIndex As Long
    Dim commissionIndex As Long
    Dim essIndex As Long
    Dim lgcIndex As Long
    Dim marketFeesIndex As Long
    Dim networkIndex As Long
    Dim retailMarginIndex As Long
    Dim revenueIndex As Long
    Dim stcIndex As Long
    Dim wholesaleEnergyIndex As Long

    nmiIndex = 1
    capacityIndex = 2
    portfolioIndex = 3
    statusIndex = 4
    associationIndex = 5
    agreementIndex = 6
    commissionIndex = 7
    essIndex = 12
    lgcIndex = 17
    marketFeesIndex = 22
    networkIndex = 27
    retailMarginIndex = 32
    revenueIndex = 37
    stcIndex = 42
    wholesaleEnergyIndex = 47
    ' Initialize JSON string
    jsonData = "["
    k = 51
    ' Ensure there are row fields And column fields
    If pvt.rowFields.Count >= 6 And pvt.ColumnFields.Count >= 1 Then
        rowFieldNmi = pvt.rowFields(1).Name
        colFieldName = pvt.ColumnFields(1).Name
        lastRow = pvt.DataBodyRange.Rows.Count + pvt.DataBodyRange.Row - 1
        ' Loop through row fields And column fields
        For i = 1 To lastRow Step k

            Set rowItemNmi = pvt.DataBodyRange.Cells( nmiIndex, 0)
            If rowItemNmi = "Grand Total" Then
                Exit For 'End
            End If
            Set rowItemCapacity = pvt.DataBodyRange.Cells( capacityIndex, 0)
            Set rowItemCommission = pvt.DataBodyRange.Cells ( commissionIndex, 0)
            Set rowItemEss = pvt.DataBodyRange.Cells ( essIndex, 0)
            Set rowItemLgc = pvt.DataBodyRange.Cells ( lgcIndex, 0)
            Set rowItemMarketFees = pvt.DataBodyRange.Cells ( marketFeesIndex, 0)
            Set rowItemNetwork = pvt.DataBodyRange.Cells ( networkIndex, 0)
            Set rowItemRetailMargin = pvt.DataBodyRange.Cells ( retailMarginIndex, 0)
            Set rowItemRevenue = pvt.DataBodyRange.Cells ( revenueIndex, 0)
            Set rowItemStc = pvt.DataBodyRange.Cells ( stcIndex, 0)
            Set rowItemWholesaleEnergy = pvt.DataBodyRange.Cells ( wholesaleEnergyIndex, 0)

            Set rowItemPortfolio = pvt.DataBodyRange.Cells(portfolioIndex, 0)
            Set rowItemStatus = pvt.DataBodyRange.Cells( statusIndex, 0)
            Set rowItemAssociation = pvt.DataBodyRange.Cells( associationIndex, 0)
            Set rowItemAgreement = pvt.DataBodyRange.Cells( agreementIndex, 0)


            For j = 1 To pvt.ColumnFields(1).PivotItems.Count
                dataFieldName = pvt.DataFields(j).Name

                ' On Error Resume Next ' Ignore error If GetPivotData fails
                valueNmi = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , colFieldName, dataFieldName)
                valueCapacity = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" , rowItemCapacity , colFieldName, dataFieldName)
                valueCommission = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" , rowItemCommission , colFieldName, dataFieldName)
                valueEss = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" ,rowItemEss , colFieldName, dataFieldName)
                valueLgc = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" , rowItemLgc, colFieldName, dataFieldName)
                valueMarketFees = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" ,rowItemMarketFees , colFieldName, dataFieldName)
                valueNetwork = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" ,rowItemNetwork , colFieldName, dataFieldName)
                valueRetailMargin = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" ,rowItemRetailMargin , colFieldName, dataFieldName)
                valueRevenue = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" , rowItemRevenue, colFieldName, dataFieldName)
                valueStc = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" ,rowItemStc , colFieldName, dataFieldName)
                valueWholesaleEnergy = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi , "Type" , rowItemWholesaleEnergy, colFieldName, dataFieldName)

                cellPortfolio = rowItemPortfolio
                cellStatus = rowItemStatus
                cellAssociation = rowItemAssociation
                cellAgreement = rowItemAgreement
                ' On Error Goto 0 ' Reset error handling

                ' Check If cellValue is Not an error And add To JSON If valid
                If Not IsError(cellValue) Then
                    jsonData = jsonData & "{" & _
                    """nmi"":""" & rowItemNmi & """," & _
                    """margin"":""" & dataFieldName & """," & _
                    """value"":" & valueNmi & "," & _
                    """type"":[" & _
                    "{""name"":""capacity"",""value"":" & valueCapacity & "}," & _
                    "{""name"":""commission"",""value"":" & valueCommission & "}," & _
                    "{""name"":""ess"",""value"":" & valueEss & "}," & _
                    "{""name"":""lgc"",""value"":" & valueLgc & "}," & _
                    "{""name"":""marketFees"",""value"":" & valueMarketFees & "}," & _
                    "{""name"":""network"",""value"":" & valueNetwork & "}," & _
                    "{""name"":""retailMargin"",""value"":" & valueRetailMargin & "}," & _
                    "{""name"":""revenue"",""value"":" & valueRevenue & "}," & _
                    "{""name"":""stc"",""value"":" & valueStc & "}," & _
                    "{""name"":""wholesaleEnergy"",""value"":" & valueWholesaleEnergy & "}" & _
                    "]," & _
                    """portfolio"":""" & cellPortfolio & """," & _
                    """status"":""" & cellStatus & """," & _
                    """association"":""" & cellAssociation & """," & _
                    """agreement"":""" & cellAgreement & """" & _
                    "},"
                End If                
            Next j

            nmiIndex = nmiIndex + k 
            portfolioIndex = portfolioIndex + k
            statusIndex = statusIndex + k
            associationIndex = associationIndex + k
            agreementIndex = agreementIndex + k

            capacityIndex = capacityIndex + k            
            commissionIndex = commissionIndex + k
            essIndex = essIndex + k
            lgcIndex = lgcIndex + k
            marketFeesIndex = marketFeesIndex + k
            networkIndex = networkIndex + k
            retailMarginIndex = retailMarginIndex + k
            revenueIndex = revenueIndex + k
            stcIndex = stcIndex + k
            wholesaleEnergyIndex = wholesaleEnergyIndex + k
        Next i

    End If

    ' Remove the trailing comma And close the JSON array
    If Len(jsonData) > 1 Then
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove last comma
    End If
    jsonData = jsonData & "]"

    ' Return the JSON string
    ExportPivotTableToJSON = jsonData
End Function


