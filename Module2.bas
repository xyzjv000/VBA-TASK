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
    Dim asssociation As String
    Dim agreement As String

    Set wsConfig = Sheets("Configurations")
    targetSheetName = wsConfig.Range("B2").Value
    startingPoint = wsConfig.Range("B4").Value
    endPoint = wsConfig.Range("B3").Value
    checksum = wsConfig.Range("B5").Value
    portfolio = wsConfig.Range("B10").Value
    statusCell = wsConfig.Range("B11").Value
    asssociation = wsConfig.Range("B12").Value
    agreement = wsConfig.Range("B13").Value

    ' Set the source sheet
    Set sourceSheet = Sheets(targetSheetName)

    ' Find the last row in column B
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, endPoint).End(xlUp).Row

    ' Define the range to copy
    Set rng = sourceSheet.Range(startingPoint & ":" & checksum & lastRow)

    ' Create a new sheet and paste data
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = setSheetName
    rng.Copy Destination:=newSheet.Range("A1")

    sourceSheet.Range(portfolio & "14:" & portfolio & lastRow).Copy

    newSheet.Range("E1").Select
    ActiveSheet.Paste

    sourceSheet.Range(statusCell & "14:" & statusCell & lastRow).Copy

    newSheet.Range("F1").Select
    ActiveSheet.Paste

    sourceSheet.Range(asssociation & "14:" & asssociation & lastRow).Copy

    newSheet.Range("G1").Select
    ActiveSheet.Paste

    sourceSheet.Range(agreement & "14:" & agreement & lastRow).Copy

    newSheet.Range("H1").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$H$" &lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6) _
        , Header:=xlNo
    Range("A1").Select
    Range(Selection, "H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
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
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim columnsArray As Variant
    Dim i As Integer
    Dim colLetter As String
    Dim formulaString As String
    Dim valueToPass As Variant
    Dim criteria As Variant
    Dim wsConfig As Worksheet
    Dim targetSheetName As String
    Dim analysisReference As String
    Dim marginStartingCell As String
    Dim nmi As String    
    Dim j As Long

    GenerateColumnAddressesArray
    Set wsConfig = Sheets("Configurations") 
    targetSheetName = wsConfig.Range("B2").Value
    analysisReference = wsConfig.Range("B7").Value
    nmi = wsConfig.Range("B3").Value
    marginStartingCell = wsConfig.Range("B6").Value
    valueToPass = tableReference & " Test"
    getData valueToPass
    Select Case tableReference 
        Case "Ancillary Services" 
            criteria = "ESS"
        Case Else 
            criteria = tableReference
    End Select
    

    ' Set your worksheet
    Set ws = ActiveSheet

    ' Specify the starting row
    startRow = 5 ' Change this to your desired starting row

    ' Find the last row with data in column A or B (whichever you expect to have the last row)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Apply the concatenation formula from startRow to lastRow in column C
    ws.Range("C" & startRow & ":C" & lastRow).Formula = "=A" & startRow & "&B" & startRow
    ws.Range("D" & startRow & ":D" & lastRow).Formula = "=IF(C" & startRow & "="""","""","""& Replace(criteria, " ", "") &""")"

    ' Array of columns to apply the SUMIFS formula
    columnsArray = columnAddressesArray

    ' Loop through each column in the array and apply the formula in each column of column E to the last row
    For i = LBound(columnsArray) To UBound(columnsArray)
        colLetter = columnsArray(i)
        Select Case colLetter
            Case "TAM"
                formulaString = "=E" & startRow & "+I" & startRow & "+M" & startRow & "+Q" & startRow & "+U" & startRow & "+Y" & startRow & "+AC" & startRow & "+AG" & startRow & "+AK" & startRow & "+AO" & startRow & "+AS" & startRow & "+AW" & startRow
            Case "TPOE90"
                formulaString = "=F" & startRow & "+J" & startRow & "+N" & startRow & "+R" & startRow & "+V" & startRow & "+Z" & startRow & "+AD" & startRow & "+AH" & startRow & "+AL" & startRow & "+AP" & startRow & "+AT" & startRow & "+AX" & startRow
            Case "TPOE50"
                formulaString = "=G" & startRow & "+K" & startRow & "+O" & startRow & "+S" & startRow & "+W" & startRow & "+AA" & startRow & "+AE" & startRow & "+AI" & startRow & "+AM" & startRow & "+AQ" & startRow & "+AU" & startRow & "+AY" & startRow
            Case "TPOE10"
                formulaString = "=H" & startRow & "+L" & startRow & "+P" & startRow & "+T" & startRow & "+X" & startRow & "+AB" & startRow & "+AF" & startRow & "+AJ" & startRow & "+AN" & startRow & "+AR" & startRow & "+AV" & startRow & "+AZ" & startRow
            Case Else
                formulaString = "=SUMIFS('"& targetSheetName &"'!" & colLetter & ":" & colLetter & _
                ", '"& targetSheetName &"'!$"& nmi &":$"& nmi &", A" & startRow & _
                ", '"& targetSheetName &"'!$"& analysisReference &":$"& analysisReference &", """ & tableReference  & """)"

        End Select
        
        ' Apply the formula to the appropriate range
        ws.Range(marginStartingCell & startRow & ":" & marginStartingCell & lastRow).Offset(0, i).Formula = formulaString
    Next i
    Range("E5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0"
    Range("A5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.UsedRange.Value = ws.UsedRange.Value
    Selection.Copy

    ReplaceOriginalTables criteria
    Application.DisplayAlerts = False ' Disable the confirmation prompt
    Sheets(valueToPass).Delete
    Application.DisplayAlerts = True  ' Re-enable the confirmation prompt    
End Sub

Public  Sub ReplaceOriginalTables(tableReference As Variant)
    Dim sheetNames As Variant
    Dim i As Integer
    ' Set the sheet where the data will be combined
    Sheets(tableReference).Select
    Range("A5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Public  Sub PopulateCombineTable()
    Dim combinedSheet As Worksheet
    Dim sheetNames As Variant
    Dim pasteRow As Integer
    Dim i As Integer
    
    sheetNames = Array("Retail Margin", "Network", "Capacity", "Wholesale Energy", _
                       "Market Fees", "ESS", "LGC", "STC", "Commission", "Revenue")
                       
    ' Set the sheet where the data will be combined
    Set combinedSheet = Sheets("Combined")

    ' Start pasting data at row 5 in the Combined sheet
    pasteRow = 5

    ' Loop through each sheet in the array
    For i = LBound(sheetNames) To UBound(sheetNames)
        Sheets(sheetNames(i)).Select
        Range("A5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        combinedSheet.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteValues
        
        ' Update the pasteRow to the next row after the last pasted data
        pasteRow = pasteRow + Selection.Rows.Count
    Next i
    Application.CutCopyMode = False
    RefreshAllPivotTables
    Sheets("Run Sheet").Select
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

    Set wsConfig = Sheets("Configurations")
    ' AP
    startRef = wsConfig.Range("B8").Value
    ' EC
    endRef = wsConfig.Range("B9").Value 
    
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
    
    ' Resize the array to the exact number of addresses
    ReDim Preserve addressArray(1 To addressCount + 4)
    addressArray(addressCount + 1) = "TAM"
    addressArray(addressCount + 2) = "TPOE90"
    addressArray(addressCount + 3) = "TPOE50"
    addressArray(addressCount + 4) = "TPOE10"
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
