Private Sub Workbook_Open()

    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets

        On Error Resume Next

        ws.Protect Password:="YOUR_PASSWORD", UserInterfaceOnly:=True

        On Error GoTo 0

    Next ws

    Application.ScreenUpdating = True

End Sub