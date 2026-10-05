Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In ThisWorkbook.Worksheets

        LockPopulatedCells ws.Range("C3:C115")
        LockPopulatedCells ws.Range("E3:E115")
        LockPopulatedCells ws.Range("B117:B119")
        LockPopulatedCells ws.Range("B121:B123")
        LockPopulatedCells ws.Range("B125")
        LockPopulatedCells ws.Range("B127")
        LockPopulatedCells ws.Range("B129:B130")

    Next ws

CleanExit:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "An error occurred while automatically locking cells." & vbCrLf &
           "Sheet: " & ws.Name & vbCrLf &
           "Error: " & Err.Number & vbCrLf &
           Err.Description,
           vbExclamation, "Automation Error"

End Sub


Private Sub LockPopulatedCells(ByVal rng As Range)

    Dim cell As Range

    For Each cell In rng

        If Not cell.MergeCells Then

            If cell.Locked = False And Len(cell.Value) > 0 Then
                cell.Locked = True
            End If

        End If

    Next cell

End Sub