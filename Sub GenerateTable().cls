Sub GenerateInsertData()
    Dim wsConfig As Worksheet
    Dim wsDestination As Worksheet
    Dim wsData As Worksheet
    Dim sourceSheetName As String
    Dim destinationSheetName As String
    Dim startColumn As String
    Dim currentColumn As String
    Dim i As Integer
    Dim nmiCount As Integer
    Dim batch As Integer
    Dim months As Integer

    Dim lastRow As Long
    Dim startRow As Long
    Dim destinationRow As Long
    Dim sourceLastRow As Long
    Dim rowValues As Variant
    Dim combinedString As String
    Dim cellValue As Variant
    Dim joinedText As String ' For concatenating all column N values

    Set wsConfig = Sheets("Config")
    sourceSheetName = wsConfig.Range("B6").Value
    destinationSheetName = wsConfig.Range("B7").Value
    months = wsConfig.Range("B10").Value

    Set wsData = Sheets(sourceSheetName)
    Set wsDestination = Sheets(destinationSheetName)

    startRow = 13 ' Change this to your desired starting row
    destinationRow = 2
    ' Find the last row with data in column A or B (whichever you expect to have the last row)
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Outer loop: Repeat for 12 sets of columns
    For batch = 1 To months
        ' Set the starting column for this batch
        startColumn = GetColumnLetter(Range("BB1").Column + (batch - 1) * 8)

        For nmiCount = startRow To lastRow
            If wsData.Range("A" & nmiCount).Value = "" Then
                wsDestination.Range("A" & destinationRow).Value = wsDestination.Range("A" & destinationRow - 1).Value
            Else
                wsDestination.Range("A" & destinationRow).Value = wsData.Range("A" & nmiCount).Value
            End If

            If wsData.Range("B" & nmiCount).Value = "" Then
                wsDestination.Range("B" & destinationRow).Value = wsDestination.Range("B" & destinationRow - 1).Value
            Else
                wsDestination.Range("B" & destinationRow).Value = wsData.Range("B" & nmiCount).Value
            End If

            If wsData.Range("C" & nmiCount).Value = "" Then
                wsDestination.Range("C" & destinationRow).Value = wsDestination.Range("C" & destinationRow - 1).Value
            Else
                wsDestination.Range("C" & destinationRow).Value = wsData.Range("C" & nmiCount).Value
            End If

            wsDestination.Range("D" & destinationRow).Value = wsData.Range("AC" & nmiCount).Value

            ' Inner loop: Process 8 columns starting from `startColumn`
            For i = 1 To 8
                currentColumn = GetColumnLetter(Range(startColumn & "1").Column + i - 1)
                wsDestination.Cells(destinationRow, 5 + i - 1).Value = wsData.Range(currentColumn & nmiCount).Value
            Next i

            ' Example of additional data in column M (optional)
            wsDestination.Range("M" & destinationRow).Value = Format(wsData.Range(startColumn & "10").Value, "mmm-yy")

            destinationRow = destinationRow + 1
        Next nmiCount
    Next batch

    Call CompareAndPopulateNMIData

    destinationRow = 2
    ' Outer loop: Repeat for 12 sets of columns
    For batch = 1 To months
        ' Set the starting column for this batch
        startColumn = GetColumnLetter(Range("BB1").Column + (batch - 1) * 8)

        For nmiCount = startRow To lastRow            
            ' Generate the combined string for column N
            combinedString = "(" ' Start the combined string
            rowValues = wsDestination.Range("A" & destinationRow & ":M" & destinationRow).Value ' Get values from A to M
            
            For Each cellValue In rowValues
                If IsEmpty(cellValue) Or Trim(cellValue) = "" Or UCase(cellValue) = "N/A" Then
                    combinedString = combinedString & "NULL, "
                ElseIf IsNumeric(cellValue) Then
                    combinedString = combinedString & cellValue & ", "
                Else
                    combinedString = combinedString & "'" & cellValue & "', "
                End If
            Next cellValue

            ' Remove the trailing ", " and close the parenthesis
            If Right(combinedString, 2) = ", " Then combinedString = Left(combinedString, Len(combinedString) - 2)
            
            ' Check if it's the last row and avoid adding the trailing comma
            If nmiCount = lastRow Then
                combinedString = Left(combinedString, Len(combinedString)) ' Remove the trailing comma
                combinedString = combinedString & ")" ' Close the parenthesis
            Else
                combinedString = combinedString & ")"
            End If

            ' Set the combined string in column N
            wsDestination.Range("N" & destinationRow).Value = combinedString

            destinationRow = destinationRow + 1
        Next nmiCount
    Next batch
End Sub


Function GetColumnLetter(colNum As Long) As String
    Dim colLetter As String
    colLetter = ""
    Do While colNum > 0
        colLetter = Chr((colNum - 1) Mod 26 + 65) & colLetter
        colNum = (colNum - 1) \ 26
    Loop
    GetColumnLetter = colLetter
End Function

Sub InsertDataToAzureSQL()
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim destinationSheetName As String
    Dim lastRow As Long
    Dim i As Long, batchSize As Long, batchCount As Long
    Dim conn As Object
    Dim sql As String, values As String

    Dim connString As String
    connString = "Provider=SQLOLEDB;Data Source=dias-ae.database.windows.net;" & _
                 "Initial Catalog=dias-dev;User ID=dias_sqladmin;" & _
                 "Password=Mg8#SjcCS5L4SMY:!IvmkHuBW1FQom;Encrypt=True;TrustServerCertificate=False;"

    ' Set the Config worksheet and destination sheet name
    Set wsConfig = Sheets("Config")
    destinationSheetName = wsConfig.Range("B7").Value

    ' Set the destination worksheet
    Set ws = ThisWorkbook.Sheets(destinationSheetName)

    ' Find the last row in column N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    ' Define batch size
    batchSize = 100
    batchCount = 0

    ' Create a new connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString

    On Error GoTo ErrorHandler

    For i = 2 To lastRow
        ' Build VALUES string for the current row
        If Trim(ws.Range("N" & i).Value) <> "" Then
            If values <> "" Then values = values & ", "
            values = values & ws.Range("N" & i).Value

            ' If batch size reached or it's the last row, execute the query
            If batchCount >= batchSize - 1 Or i = lastRow Then
                sql = "INSERT INTO [dbo].[ForecastResult] " & _
                      "([Customer], [NMI], [Checksum], [AnalysisData], [AchievedMargin], " & _
                      "[Predicted90POE], [Predicted50POE], [Predicted10POE], [PredictedTM], " & _
                      "[LoadP_MWh], [LoadO_P_MWh], [LoadT_MWh], [Period]) VALUES " & values

                conn.Execute sql
                values = "" ' Reset the batch values
                batchCount = 0
            Else
                batchCount = batchCount + 1
            End If
        End If
    Next i

    MsgBox "Data Inserted Successfully into Azure SQL Database!", vbInformation

    ' Cleanup
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub


Private Sub CompareAndPopulateNMIData()
    Dim wsMargin As Worksheet
    Dim wsForecast As Worksheet
    Dim marginLastRow As Long
    Dim forecastLastRow As Long
    Dim marginNMI As String
    Dim forecastAnalysis As String
    Dim marginAnalysis As String
    Dim forecastNMI As String
    Dim marginValue As Variant
    Dim marginDate As Date
    Dim marginNextMonth As String
    Dim i As Long, j As Long, k As Long

    ' Set worksheets
    Set wsMargin = ThisWorkbook.Sheets("Margin Per NMI")
    Set wsForecast = ThisWorkbook.Sheets("Generate Forecast Import")

    ' Find last rows in both sheets
    marginLastRow = wsMargin.Cells(wsMargin.Rows.Count, "E").End(xlUp).Row
    forecastLastRow = wsForecast.Cells(wsForecast.Rows.Count, "D").End(xlUp).Row

    If IsDate(wsMargin.Range("F2").Value) Then
        marginDate = wsMargin.Range("F2").Value
        ' Calculate the Next month And format it As "mmm-yy"
        marginNextMonth = Format(DateAdd("m", 1, marginDate), "mmm-yy")
    Else
        MsgBox "Invalid date in F2 on Margin Per NMI sheet.", vbExclamation
     Exit Sub
    End If

    ' Loop through each NMI in Generate Forecast Import (D column)
    For i = 2 To forecastLastRow
        forecastNMI = wsForecast.Cells(i, "B").Value

        For j = 5 To marginLastRow
            If wsMargin.Cells(j, "D").Value <> "" Then
                marginNMI = wsMargin.Cells(j, "D").Value
            End If

            forecastAnalysis = wsForecast.Cells(i, "D").Value
            marginAnalysis = wsMargin.Cells(j, "E").Value


            If forecastNMI = marginNMI And forecastAnalysis = marginAnalysis Then

                If wsForecast.Cells(i, "M").Value = marginNextMonth Then

                    wsForecast.Cells(i, "E").Value = wsMargin.Cells(j, "F").Value
                End If
             Exit For '
            End If
        Next j
    Next i

End Sub

Sub UpdateAchievedMarginToAzureSQL()
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim destinationSheetName As String
    Dim lastRow As Long
    Dim i As Long, batchSize As Long, batchCount As Long
    Dim conn As Object
    Dim sql As String, values As String
    Dim marginDate As Date
    Dim marginNextMonth As String
    Dim achievedMarginSheet As String
    Dim connString As String
    
    ' Get the configuration values
    Set wsConfig = Sheets("Config")
    achievedMarginSheet = wsConfig.Range("B11").Value
    
    ' Set the connection string
    connString = "Provider=SQLOLEDB;Data Source=dias-ae.database.windows.net;" & _
                 "Initial Catalog=dias-dev;User ID=dias_sqladmin;" & _
                 "Password=Mg8#SjcCS5L4SMY:!IvmkHuBW1FQom;Encrypt=True;TrustServerCertificate=False;"
    
    ' Set the destination worksheet and destination sheet name
    destinationSheetName = wsConfig.Range("B7").Value
    Set ws = ThisWorkbook.Sheets(destinationSheetName)

    ' Get margin date from the "Margin Per NMI" sheet (F2)
    If IsDate(Sheets("Margin Per NMI").Range("F2").Value) Then
        marginDate = Sheets("Margin Per NMI").Range("F2").Value
        ' Calculate next month and format it as "mmm-yy"
        marginNextMonth = Format(DateAdd("m", 1, marginDate), "mmm-yy")
    Else
        MsgBox "Invalid date in F2 on Margin Per NMI sheet.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in column N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    ' Define batch size
    batchSize = 100
    batchCount = 0
    values = "" ' Initialize values string

    ' Create a new connection to the database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString

    On Error GoTo ErrorHandler

    ' Loop through all rows in the destination sheet (ws) column N
    For i = 2 To lastRow
        ' Only add the value to the batch if it matches the next month's value
        If Trim(ws.Range("M" & i).Value) <> "" And _
           ws.Range("M" & i).Value = marginNextMonth Then
            
            ' Build VALUES string for the current row
            If values <> "" Then values = values & ", "
            values = values & ws.Range("N" & i).Value

            ' If batch size reached or it's the last row, execute the query
            If batchCount >= batchSize - 1 Or i = lastRow Then
                sql = "INSERT INTO [dbo].[ForecastResult] " & _
                      "([Customer], [NMI], [Checksum], [AnalysisData], [AchievedMargin], " & _
                      "[Predicted90POE], [Predicted50POE], [Predicted10POE], [PredictedTM], " & _
                      "[LoadP_MWh], [LoadO_P_MWh], [LoadT_MWh], [Period]) VALUES " & values
                conn.Execute sql
                values = "" ' Reset the batch values
                batchCount = 0
            Else
                batchCount = batchCount + 1
            End If
        End If
    Next i

    ' If there are any remaining rows that didn't fill a full batch, insert them now
    If values <> "" Then
        sql = "INSERT INTO [dbo].[ForecastResult] " & _
              "([Customer], [NMI], [Checksum], [AnalysisData], [AchievedMargin], " & _
              "[Predicted90POE], [Predicted50POE], [Predicted10POE], [PredictedTM], " & _
              "[LoadP_MWh], [LoadO_P_MWh], [LoadT_MWh], [Period]) VALUES " & values
        conn.Execute sql
    End If

    MsgBox "New Data Inserted Successfully!", vbInformation

    ' Cleanup
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub
