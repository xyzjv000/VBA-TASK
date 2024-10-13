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
    Dim combinedJson As String
    Dim summaryJson As String
    Dim nmiJson As String
    Dim combinedFullJson As String
    Dim templateFilePath As String
    Dim cssFilePath As String
    Dim jsFilePath As String
    Dim templateFileNumber As Integer
    Dim jsonResult3 As String
    Dim jsonResult4 As String
    Dim versionData As String

    ' JSON extraction from your Pivot Tables
    jsonResult3 = ExtractFilteredDataToJSONArrayMargins("Exported Data", "RetailMarginPivot")
    jsonResult4 = ExportPivotTableToJSON("Exported Data", "CombinedDataPivot")

    ' File paths
    templateFilePath = GetCurrentExcelDirectory & Application.PathSeparator & "Exports" & Application.PathSeparator & "HTML_Template.html"
    cssFilePath = GetCurrentExcelDirectory & Application.PathSeparator & "Exports" & Application.PathSeparator & "styles.css"
    jsFilePath = GetCurrentExcelDirectory & Application.PathSeparator & "Exports" & Application.PathSeparator & "script.js"

    ' Read HTML template from file
    templateFileNumber = FreeFile
    Open templateFilePath For Input As #templateFileNumber
    htmlTemplate = Input$(LOF(templateFileNumber), templateFileNumber)
    Close #templateFileNumber

    ' Replace placeholders in the HTML template With actual data
    nmiJson = jsonResult3
    combinedFullJson = jsonResult4
    versionData = GetVersionData()

    ' htmlContent = Replace(htmlTemplate, "{{combinedJson}}", combinedJson)
    ' htmlContent = Replace(htmlContent, "{{summaryJson}}", summaryJson)
    htmlContent = Replace(htmlTemplate, "{{versionData}}", versionData)
    htmlContent = Replace(htmlContent, "{{nmiJson}}", nmiJson)
    htmlContent = Replace(htmlContent, "{{combinedFullJson}}", combinedFullJson)

    ' Replace CSS And JS placeholders With file paths
    htmlContent = Replace(htmlContent, "{{cssFilePath}}", cssFilePath)
    htmlContent = Replace(htmlContent, "{{jsFilePath}}", jsFilePath)

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

' USED
Function ExtractFilteredDataToJSONArrayMargins(sheetName As String, pivotTableName As String) As String
    Dim ws As Worksheet
    Dim pvt As PivotTable
    Dim jsonData As String
    Dim jsonDataItem As String
    Dim rowItem As PivotItem
    Dim rowItemNmi As Variant
    Dim rowItemPortfolio As Variant
    Dim rowItemStatus As Variant
    Dim rowItemAssociation As Variant
    Dim rowItemAgreement As Variant
    Dim colItem As PivotItem
    Dim valueNmi As Variant
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim colFieldName As String
    Dim dataFieldName As String

    ' Set worksheet And pivot table
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set pvt = ws.PivotTables(pivotTableName)

    ' Index setup
    Dim nmiIndex As Long
    Dim portfolioIndex As Long
    Dim statusIndex As Long
    Dim associationIndex As Long
    Dim agreementIndex As Long

    nmiIndex = 1
    statusIndex = 2
    portfolioIndex = 3
    associationIndex = 4
    agreementIndex = 5
    ' Initialize JSON string
    jsonData = "["

    k = 5
    ' Ensure row And column fields exist
    If pvt.rowFields.Count >= 5 And pvt.ColumnFields.Count >= 1 Then
        rowFieldNmi = pvt.rowFields(1).name
        colFieldName = pvt.ColumnFields(1).name
        lastRow = pvt.DataBodyRange.Rows.Count + pvt.DataBodyRange.row - 1

        ' Loop through row fields And column fields
        For i = 1 To lastRow Step k
            Set rowItemNmi = pvt.DataBodyRange.Cells(nmiIndex, 0)
            If rowItemNmi = "Grand Total" Then
             Exit For
            End If
            If rowItemNmi = "(blank)" Then
             Exit For
            End If
            Set rowItemPortfolio = pvt.DataBodyRange.Cells(portfolioIndex, 0)
            Set rowItemStatus = pvt.DataBodyRange.Cells(statusIndex, 0)
            Set rowItemAssociation = pvt.DataBodyRange.Cells(associationIndex, 0)
            Set rowItemAgreement = pvt.DataBodyRange.Cells(agreementIndex, 0)

            ' Loop through column items
            For j = 1 To pvt.ColumnFields(1).PivotItems.Count - 3
                dataFieldName = pvt.DataBodyRange.Cells(0, j)
                If dataFieldName = "Grand Total" Then
                 Exit For
                End If
                If dataFieldName <> "(blank)" Then
                    valueNmi = pvt.DataBodyRange.Cells(nmiIndex, j)
                    If Not IsError(cellValue) Then
                        jsonDataItem = jsonDataItem & _
                        "{""margin"":""" & dataFieldName & """," & _
                        """value"":" & valueNmi & "},"
                    End If
                End If                
            Next j
            If Len(jsonDataItem) > 1 Then
                jsonDataItem = Left(jsonDataItem, Len(jsonDataItem) - 1) ' Remove last comma
            End If
            jsonData = jsonData & "{" & _
            """nmi"":""" & rowItemNmi & """," & _
            """data"":[" & _
            jsonDataItem & _
            "]," & _
            """portfolio"":""" & rowItemPortfolio & """," & _
            """status"":""" & rowItemStatus & """," & _
            """association"":""" & rowItemAssociation & """," & _
            """agreement"":""" & rowItemAgreement & """" & _
            "},"
            jsonDataItem = ""
            ' Increment indices
            nmiIndex = nmiIndex + k
            portfolioIndex = portfolioIndex + k
            statusIndex = statusIndex + k
            associationIndex = associationIndex + k
            agreementIndex = agreementIndex + k
        Next i
    End If

    ' Remove trailing comma And close JSON array
    If Len(jsonData) > 1 Then
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove last comma
    End If
    jsonData = jsonData & "]"

    ' Return the JSON string
    ExtractFilteredDataToJSONArrayMargins = jsonData
End Function

' USED
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

' USED
Function ExportPivotTableToJSON(sheetName As String, pivotTableName As String) As String
    Dim ws As Worksheet
    Dim pvt As PivotTable
    Dim jsonData As String
    Dim jsonDataItem As String
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
    Dim rowItemSecurityDepositInterest As Variant
    Dim rowItemSecurityDeposit As Variant
    Dim rowItemRoc As Variant
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
    Dim valueSecurityDepositInterest As Variant
    Dim valueSecurityDeposit As Variant
    Dim valueRoc As Variant
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim colFieldName As String
    Dim dataFieldName As String

    ' Set worksheet And pivot table
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set pvt = ws.PivotTables(pivotTableName)

    ' Index setup
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
    Dim securityDepositInterestIndex As Long
    Dim securityDepositIndex As Long
    Dim rocIndex As Long

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
    securityDepositInterestIndex = 52
    securityDepositIndex = 57
    rocIndex = 62

    ' Initialize JSON string
    jsonData = "["

    k = 66
    ' Ensure row And column fields exist
    If pvt.rowFields.Count >= 6 And pvt.ColumnFields.Count >= 1 Then
        rowFieldNmi = pvt.rowFields(1).name
        colFieldName = pvt.ColumnFields(1).name
        lastRow = pvt.DataBodyRange.Rows.Count + pvt.DataBodyRange.row - 1

        ' Loop through row fields And column fields
        For i = 1 To lastRow Step k
            Set rowItemNmi = pvt.DataBodyRange.Cells(nmiIndex, 0)
            If rowItemNmi = "Grand Total" Then
             Exit For
            End If
            Set rowItemCapacity = pvt.DataBodyRange.Cells(capacityIndex, 0)
            Set rowItemCommission = pvt.DataBodyRange.Cells(commissionIndex, 0)
            Set rowItemEss = pvt.DataBodyRange.Cells(essIndex, 0)
            Set rowItemLgc = pvt.DataBodyRange.Cells(lgcIndex, 0)
            Set rowItemMarketFees = pvt.DataBodyRange.Cells(marketFeesIndex, 0)
            Set rowItemNetwork = pvt.DataBodyRange.Cells(networkIndex, 0)
            Set rowItemRetailMargin = pvt.DataBodyRange.Cells(retailMarginIndex, 0)
            Set rowItemRevenue = pvt.DataBodyRange.Cells(revenueIndex, 0)
            Set rowItemStc = pvt.DataBodyRange.Cells(stcIndex, 0)
            Set rowItemWholesaleEnergy = pvt.DataBodyRange.Cells(wholesaleEnergyIndex, 0)
            Set rowItemSecurityDepositInterest = pvt.DataBodyRange.Cells(securityDepositInterestIndex, 0)
            Set rowItemSecurityDeposit = pvt.DataBodyRange.Cells(securityDepositIndex, 0)
            Set rowItemRoc = pvt.DataBodyRange.Cells(rocIndex, 0)

            Set rowItemPortfolio = pvt.DataBodyRange.Cells(portfolioIndex, 0)
            Set rowItemStatus = pvt.DataBodyRange.Cells(statusIndex, 0)
            Set rowItemAssociation = pvt.DataBodyRange.Cells(associationIndex, 0)
            Set rowItemAgreement = pvt.DataBodyRange.Cells(agreementIndex, 0)

            ' Loop through column items
            For j = 1 To pvt.ColumnFields(1).PivotItems.Count
                dataFieldName = pvt.DataFields(j).name

                ' On Error Resume Next ' Ignore error If GetPivotData fails
                valueNmi = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, colFieldName, dataFieldName)
                valueCapacity = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemCapacity, colFieldName, dataFieldName)
                valueCommission = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemCommission, colFieldName, dataFieldName)
                valueEss = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemEss, colFieldName, dataFieldName)
                valueLgc = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemLgc, colFieldName, dataFieldName)
                valueMarketFees = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemMarketFees, colFieldName, dataFieldName)
                valueNetwork = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemNetwork, colFieldName, dataFieldName)
                valueRetailMargin = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemRetailMargin, colFieldName, dataFieldName)
                valueRevenue = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemRevenue, colFieldName, dataFieldName)
                valueStc = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemStc, colFieldName, dataFieldName)
                valueWholesaleEnergy = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemWholesaleEnergy, colFieldName, dataFieldName)
                valueSecurityDepositInterest = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemSecurityDepositInterest, colFieldName, dataFieldName)
                valueSecurityDeposit = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemSecurityDeposit, colFieldName, dataFieldName)
                valueRoc = pvt.GetPivotData(dataFieldName, rowFieldNmi, rowItemNmi, "Type", rowItemRoc, colFieldName, dataFieldName)

                cellPortfolio = rowItemPortfolio
                cellStatus = rowItemStatus
                cellAssociation = rowItemAssociation
                cellAgreement = rowItemAgreement
                ' On Error Goto 0 ' Reset error handling

                ' Check If cellValue is Not an error And add To JSON If valid
                If Not IsError(cellValue) Then
                    jsonDataItem = jsonDataItem & _
                    "{""margin"":""" & dataFieldName & """," & _
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
                    "{""name"":""wholesaleEnergy"",""value"":" & valueWholesaleEnergy & "}," & _
                    "{""name"":""securityDepositInterest"",""value"":" & valueSecurityDepositInterest & "}," & _
                    "{""name"":""securityDeposit"",""value"":" & valueSecurityDeposit & "}," & _
                    "{""name"":""roc"",""value"":" & valueRoc & "}" & _
                    "]},"
                End If
            Next j

            jsonData = jsonData & "{" & _
            """nmi"":""" & rowItemNmi & """," & _
            """data"":[" & _
            jsonDataItem & _
            "]," & _
            """portfolio"":""" & rowItemPortfolio & """," & _
            """status"":""" & rowItemStatus & """," & _
            """association"":""" & rowItemAssociation & """," & _
            """agreement"":""" & rowItemAgreement & """" & _
            "},"
            jsonDataItem = ""
            ' Increment indices
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
            securityDepositInterestIndex = securityDepositInterestIndex + k
            securityDepositIndex = securityDepositIndex + k
            rocIndex = rocIndex + k
        Next i
    End If

    ' Remove trailing comma And close JSON array
    If Len(jsonData) > 1 Then
        jsonData = Left(jsonData, Len(jsonData) - 1) ' Remove last comma
    End If
    jsonData = jsonData & "]"

    ' Return the JSON string
    ExportPivotTableToJSON = jsonData
End Function