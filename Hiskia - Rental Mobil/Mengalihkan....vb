Public Class Mengalihkan

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim go As New Menu_Utama
        ProgressBar1.Value += 25
        If ProgressBar1.Value = 100 Then
            Timer1.Dispose()
            Me.Close()
            go.Show()
        End If
    End Sub
End Class