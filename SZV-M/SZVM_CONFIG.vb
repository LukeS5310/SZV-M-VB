Public Class SZVM_CONFIG
    Private Sub BTN_CLEAN_ALL_Click(sender As Object, e As EventArgs) Handles BTN_CLEAN_ALL.Click
        CleanUP(CleanType.ALL)
        MsgBox("Очистка базы завершена")
    End Sub
End Class