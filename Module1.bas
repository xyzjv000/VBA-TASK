Sub RunInBackgroundGenerate()
    ' Turn off screen updating, automatic calculation, and events
    Dim response As VbMsgBoxResult

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup
    Call Module2.GenerateTables
Cleanup:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub RunInBackgroundDelete()
    ' Turn off screen updating, automatic calculation, and events
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup
    Call Module2.DeleteData
Cleanup:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub ClearFilter()
    ' Turn off screen updating, automatic calculation, and events
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup
    Call Module2.ClearAllSlicerFilters
Cleanup:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub