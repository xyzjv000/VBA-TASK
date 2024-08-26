Option Private Module
Option Explicit
Dim columnAddressesArray() As String

Public Sub getData(setSheetName As Variant)
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim targetSheetName As String
    Dim startingPoint As String
    Dim endPoint As String
    Dim checksum As String
    Dim portfolio As String
    Dim statusCell As String
    Dim association As String
    Dim agreement As String
    Dim copiedSheetRange As String

    On Error GoTo ErrorHandler

    ' Set the configuration worksheet
    Set wsConfig = Sheets("Configurations")
    
    ' Retrieve configuration values
    targetSheetName = wsConfig.Range("B2").Value
    startingPoint = wsConfig.Range("B4").Value
    endPoint = wsConfig.Range("B3").Value
    checksum = wsConfig.Range("B5").Value
    portfolio = wsConfig.Range("B10").Value
    statusCell = wsConfig.Range("B11").Value
    association = wsConfig.Range("B12").Value
    agreement = wsConfig.Range("B13").Value
    copiedSheetRange = wsConfig.Range("B26").Value

    ' Set the source sheet
    Set sourceSheet = Sheets(targetSheetName)

    ' Find the last row in the specified column
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, endPoint).End(xlUp).Row

    ' Define the range to copy
    Set rng = sourceSheet.Range(startingPoint & ":" & checksum & lastRow)

    ' Create a new sheet and paste data
    Set newSheet = Sheets(setSheetName)
    newSheet.Activate
    ' Copy and paste data ranges
    rng.Copy Destination:=newSheet.Range("A5")
    ' Remove duplicates
    newSheet.Range(copiedSheetRange & lastRow).RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo

    ' Clean up
    Application.CutCopyMode = False

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub


' GENERATE DATA
Public Sub GenerateTables()
    Dim products As Variant
    Dim i As Integer
    Dim response As VbMsgBoxResult

    ' Display a message box with Yes and No options
    response = MsgBox("Do you want to proceed with this action?", vbYesNo + vbQuestion, "Confirm Action")
    
    ' Check the user's response
    If response = vbYes Then
        ' Add the code to be executed if the user clicks Yes
            ' products = Array("Retail Margin")
            products = Array("Retail Margin", "Network", "Capacity", "Wholesale Energy", "Market Fees", "Ancillary Services", "LGC", "STC", "Commission", "Revenue")

            For i = LBound(products) To UBound(products)
                TableTemplate products(i)
            Next i
        
            Call PopulateCombineTable
        MsgBox "The action was completed successfully.", vbInformation, "Success"
    Else
        ' Optionally, add code to be executed if the user clicks No
        MsgBox "Action cancelled.", vbInformation, "Cancelled"
    End If
End Sub

Public Sub TableTemplate(tableReference As Variant)
    ' Worksheet Variables
    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim sourceSheet As Worksheet
    ' String Variables
    Dim TPStartColumn As String
    Dim TAMStartColumn As String
    Dim TP90StartColumn As String
    Dim TP50StartColumn As String
    Dim TP10StartColumn As String
    Dim NextTPStartColumn As String
    Dim NextTAMStartColumn As String
    Dim NextTP90StartColumn As String
    Dim NextTP50StartColumn As String
    Dim NextTP10StartColumn As String
    Dim targetSheetName As String
    Dim analysisReference As String
    Dim marginStartingCell As String
    Dim nmi As String
    Dim formulaString As String
    Dim achievedCell As String
    
    ' Variant Variables
    Dim columnsArray As Variant
    Dim criteria As Variant
    
    ' Long Variables
    Dim lastRow As Long
    Dim startRow As Long
    Dim sourceLastRow As Long
    
    ' Integer Variables
    Dim i As Integer
    Dim colLetter As String


    GenerateColumnAddressesArray
    Set wsConfig = Sheets("Configurations") 
    targetSheetName = wsConfig.Range("B2").Value
    analysisReference = wsConfig.Range("B7").Value
    nmi = wsConfig.Range("B3").Value
    marginStartingCell = wsConfig.Range("B6").Value
    TPStartColumn = wsConfig.Range("B14").Value
    TAMStartColumn = wsConfig.Range("B15").Value
    TP90StartColumn = wsConfig.Range("B16").Value
    TP50StartColumn = wsConfig.Range("B17").Value
    TP10StartColumn = wsConfig.Range("B18").Value
    NextTPStartColumn = wsConfig.Range("B20").Value
    NextTAMStartColumn = wsConfig.Range("B21").Value
    NextTP90StartColumn = wsConfig.Range("B22").Value
    NextTP50StartColumn = wsConfig.Range("B23").Value
    NextTP10StartColumn = wsConfig.Range("B24").Value
    achievedCell = wsConfig.Range("B25").Value

    Set sourceSheet = Sheets(targetSheetName)
    sourceLastRow = sourceSheet.Cells(sourceSheet.Rows.Count, nmi).End(xlUp).Row

    Select Case tableReference 
        Case "Ancillary Services" 
            criteria = "ESS"
        Case Else 
            criteria = tableReference
    End Select
    getData criteria

    ' Set your worksheet
    Set ws = ActiveSheet

    ' Specify the starting row
    startRow = 5 ' Change this to your desired starting row

    ' Find the last row with data in column A or B (whichever you expect to have the last row)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("C" & startRow & ":C" & lastRow).Formula = "=A" & startRow & "&B" & startRow
    ws.Range("D" & startRow & ":D" & lastRow).Formula = "=IF(C" & startRow & "="""","""","""& Replace(criteria, " ", "") &""")"
    ws.Range("E" & startRow & ":E" & lastRow).Formula = "=IF(A" & startRow & "="""","""",VLOOKUP($A" & startRow & ",'" & targetSheetName & "'!$D$13:$M$"& sourceLastRow &",5,FALSE))"
    ws.Range("F" & startRow & ":F" & lastRow).Formula = "=IF(A" & startRow & "="""","""",VLOOKUP($A" & startRow & ",'" & targetSheetName & "'!$D$13:$M$"& sourceLastRow &",6,FALSE))"
    ws.Range("G" & startRow & ":G" & lastRow).Formula = "=IF(A" & startRow & "="""","""",VLOOKUP($A" & startRow & ",'" & targetSheetName & "'!$D$13:$M$"& sourceLastRow &",7,FALSE))"
    ws.Range("H" & startRow & ":H" & lastRow).Formula = _
    "=IF(A" & startRow & "="""","""",IF(LEFT(VLOOKUP($A" & startRow & ",'" & targetSheetName & "'!$D$13:$M$"& sourceLastRow &",10,FALSE),9)=""Unbundled"",""Unbundled"",""Bundled""))"

    ' Array of columns to apply the SUMIFS formula
    columnsArray = columnAddressesArray

    ' Loop through each column in the array and apply the formula in each column of column E to the last row
    For i = LBound(columnsArray) To UBound(columnsArray)
        colLetter = columnsArray(i)
        Select Case colLetter
            Case "Achieved"                                
                formulaString = "=INDEX('" & targetSheetName & "'!" & achievedCell & ":" & achievedCell & ", MATCH(1,('" & targetSheetName & "'!" & analysisReference & ":" & analysisReference & "=""" & tableReference  & """)*('" & targetSheetName & "'!D:D=A" & startRow & "), 0))"
            Case "TP"
                formulaString =  "=SUM(" & GenerateColumnSequence(TPStartColumn, startRow) & ")"
            Case "TAM"
                formulaString =  "=SUM(" & GenerateColumnSequence(TAMStartColumn, startRow) & ")"
            Case "TPOE90"
                formulaString =  "=SUM(" & GenerateColumnSequence(TP90StartColumn, startRow) & ")"
            Case "TPOE50"
                formulaString =  "=SUM(" & GenerateColumnSequence(TP50StartColumn, startRow) & ")"
            Case "TPOE10"
                formulaString =  "=SUM(" & GenerateColumnSequence(TP10StartColumn, startRow) & ")"
            Case "_TP"
                formulaString =  "=SUM(" & GenerateColumnSequence(NextTPStartColumn, startRow) & ")"
            Case "_TAM"
                formulaString =  "=SUM(" & GenerateColumnSequence(NextTAMStartColumn, startRow) & ")"
            Case "_TPOE90"
                formulaString =  "=SUM(" & GenerateColumnSequence(NextTP90StartColumn, startRow) & ")"
            Case "_TPOE50"
                formulaString =  "=SUM(" & GenerateColumnSequence(NextTP50StartColumn, startRow) & ")"
            Case "_TPOE10"
                formulaString =  "=SUM(" & GenerateColumnSequence(NextTP10StartColumn, startRow) & ")"
            Case Else
                formulaString = "=SUMIFS('"& targetSheetName &"'!" & colLetter & ":" & colLetter & _
                ", '"& targetSheetName &"'!$"& nmi &":$"& nmi &", A" & startRow & _
                ", '"& targetSheetName &"'!$"& analysisReference &":$"& analysisReference &", """ & tableReference  & """)"

        End Select
        ' Debug.Print formulaString
        ' Apply the formula to the appropriate range
        Debug.Print formulaString
        If colLetter = "Achieved" Then
            ws.Range(marginStartingCell & startRow & ":" & marginStartingCell & lastRow).Offset(0, i).Formula2 = formulaString
        Else
            ws.Range(marginStartingCell & startRow & ":" & marginStartingCell & lastRow).Offset(0, i).Formula = formulaString
        End If

    Next i
    ws.Range("E5", ws.Range("E5").End(xlToRight).End(xlDown)).NumberFormat = "0"

    Range("A5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.UsedRange.Value = ws.UsedRange.Value
    Selection.Copy

    ReplaceOriginalTables criteria
End Sub

Public  Sub ReplaceOriginalTables(tableReference As Variant)
    Dim sheetNames As Variant
    Dim i As Integer
    ' Set the sheet where the data will be combined
    Sheets(tableReference).Select
    Range("A5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Public Sub PopulateCombineTable()
    Dim combinedSheet As Worksheet
    Dim sheetNames As Variant
    Dim pasteRow As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rng As Range
    Dim i As Integer

    sheetNames = Array("Retail Margin", "Network", "Capacity", "Wholesale Energy", _
                       "Market Fees", "ESS", "LGC", "STC", "Commission", "Revenue")
                       
    ' Set the sheet where the data will be combined
    Set combinedSheet = Sheets("Combined")

    ' Start pasting data at row 5 in the Combined sheet
    pasteRow = 5

    ' Loop through each sheet in the array
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = Sheets(sheetNames(i))
        
        ' Find the last row and column with data
        With ws
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            lastCol = .Cells(5, .Columns.Count).End(xlToLeft).Column
            Set rng = .Range(.Cells(5, 1), .Cells(lastRow, lastCol))
        End With
        
        ' Copy data to combinedSheet
        rng.Copy
        combinedSheet.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteValues
        
        ' Update the pasteRow to the next row after the last pasted data
        pasteRow = pasteRow + rng.Rows.Count
    Next i
    
    Application.CutCopyMode = False
    RefreshAllPivotTables
End Sub


Public Sub RefreshAllPivotTables()
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim cht As ChartObject
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through all PivotTables in the worksheet and refresh them
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt

        For Each cht In ws.ChartObjects
            cht.Chart.Refresh
        Next cht
    Next ws
End Sub

' OPTIONAL FOR TESTING ONLY
Public Sub DeleteData()
    Dim response As VbMsgBoxResult
    Dim sheetNames As Variant
    Dim i As Integer
    ' Display a message box with Yes and No options
    response = MsgBox("Do you want to proceed with this DELETE action?", vbYesNo + vbQuestion, "Confirm Action")
    
        ' Check the user's response
        If response = vbYes Then
            ' Add the code to be executed if the user clicks Yes
                sheetNames = Array("Retail Margin", "Network", "Capacity", "Wholesale Energy", _
                           "Market Fees", "ESS", "LGC", "STC", "Commission", "Revenue", "Combined")
            
        For i = LBound(sheetNames) To UBound(sheetNames)
            Sheets(sheetNames(i)).Select
            Range("A5").Select
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            Selection.ClearContents            
        Next i
        RefreshAllPivotTables
        MsgBox "The action was completed successfully.", vbInformation, "Success"
    Else
        ' Optionally, add code to be executed if the user clicks No
        MsgBox "Action cancelled.", vbInformation, "Cancelled"
    End If
    Sheets("Run Sheet").Select
End Sub


Public Sub GenerateColumnAddressesArray()
    Dim startCol As Long
    Dim endCol As Long
    Dim colIndex As Long
    Dim i As Long
    Dim addressArray() As String
    Dim addressCount As Long
    Dim wsConfig As Worksheet
    Dim startRef As String   
    Dim endRef As String
    Dim nextYearIncluded As Boolean  

    Set wsConfig = Sheets("Configurations")
    ' AP
    startRef = wsConfig.Range("B8").Value
    ' EC
    endRef = wsConfig.Range("B9").Value 
    nextYearIncluded = wsConfig.Range("B19").Value 
    
    ' Define the starting and ending columns
    startCol = ColLetterToNum(startRef) ' Column number for "AP"
    endCol = ColLetterToNum(endRef) ' Column number for "EC"
    
    addressCount = 0 ' Initialize count of addresses
    
    ' Pre-allocate the array with a rough estimate of size
    ReDim addressArray(1 To 100) 
    
    For colIndex = startCol To endCol Step 8 ' Loop through every 8 columns
        For i = 0 To 4 ' Get addresses for the first 4 columns in each group
            If colIndex + i <= endCol Then
                addressCount = addressCount + 1
                If addressCount > UBound(addressArray) Then
                    ' Resize the array if necessary
                    ReDim Preserve addressArray(1 To addressCount * 2)
                End If
                addressArray(addressCount) = ColNumToLetter(colIndex + i) ' Convert column number to letter
            End If
        Next i
        ' Stop the outer loop if the next step would go beyond endCol
        If colIndex + 8 > endCol Then Exit For
    Next colIndex

    If nextYearIncluded Then
        ' Expand the array by 10 elements
        ReDim Preserve addressArray(1 To addressCount + 11)
        
        ' Add new values to the array
        addressArray(addressCount + 1) = "Achieved"
        addressArray(addressCount + 2) = "TAM"
        addressArray(addressCount + 3) = "TPOE90"
        addressArray(addressCount + 4) = "TPOE50"
        addressArray(addressCount + 5) = "TPOE10"
        addressArray(addressCount + 6) = "TP"
        addressArray(addressCount + 7) = "_TAM"
        addressArray(addressCount + 8) = "_TPOE90"
        addressArray(addressCount + 9) = "_TPOE50"
        addressArray(addressCount + 10) = "_TPOE10"
        addressArray(addressCount + 11) = "_TP"
    Else
        ' Expand the array by 5 elements
        ReDim Preserve addressArray(1 To addressCount + 6)
        
        ' Add new values to the array
        addressArray(addressCount + 1) = "Achieved"
        addressArray(addressCount + 2) = "TAM"
        addressArray(addressCount + 3) = "TPOE90"
        addressArray(addressCount + 4) = "TPOE50"
        addressArray(addressCount + 5) = "TPOE10"
        addressArray(addressCount + 6) = "TP"
    End If
    ' Store the array in the module-level variable
    columnAddressesArray = addressArray
End Sub

Function ColNumToLetter(colNum As Long) As String
    Dim colLetter As String
    Dim i As Long
    Dim tempCol As Long
    
    tempCol = colNum
    colLetter = ""
    
    While tempCol > 0
        i = (tempCol - 1) Mod 26
        colLetter = Chr(65 + i) & colLetter
        tempCol = (tempCol - i) \ 26
    Wend
    
    ColNumToLetter = colLetter
End Function

Function ColLetterToNum(colLetter As String) As Long
    Dim colNum As Long
    Dim i As Long
    Dim lenColLetter As Long
    
    lenColLetter = Len(colLetter)
    colNum = 0
    
    For i = 1 To lenColLetter
        colNum = colNum * 26 + (Asc(UCase(Mid(colLetter, i, 1))) - Asc("A") + 1)
    Next i
    
    ColLetterToNum = colNum
End Function

Function GenerateColumnSequence(col As String, rowNum As Long) As String
    Dim sequence As String
    Dim i As Integer
    
    ' Start at the initial column, then increment by 5 columns
    For i = 0 To 11 ' 12 columns in total (M, R, W, AB, AG, etc.)
        ' Calculate the column address by offsetting the initial column
        sequence = sequence & Range(col & rowNum).Offset(0, i * 5).Address(False, False) & ","
    Next i
    
    ' Remove the trailing comma
    sequence = Left(sequence, Len(sequence) - 1)
    
    ' Return the sequence
    GenerateColumnSequence = sequence
End Function
