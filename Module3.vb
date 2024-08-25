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
