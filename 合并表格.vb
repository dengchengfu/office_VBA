Private Sub Worksheet_Change(ByVal Target As Range)

    For Each MR In Target
        pd_name = MR.Value
        If pd_name = "value" Then
            tg_row = MR.Row
            tg_column = MR.Column
            Worksheets("Sheet1").Cells(tg_row, tg_column) = "target"
        End If
    Next
    
End Sub