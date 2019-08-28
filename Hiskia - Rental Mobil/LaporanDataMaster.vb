Public Class LaporanDataMaster

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        LaporanDataMobil.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        LaporanDataPetugas.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LaporanDataSupir.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LaporanDataPelanggan.Show()
        Me.Hide()
    End Sub
End Class