Public Class TestClient

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim oAlertEngine As New USEcology.AlertEngine
        oAlertEngine.DoWork()
    End Sub
End Class
