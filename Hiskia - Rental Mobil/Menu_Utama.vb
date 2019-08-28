Imports System.Data.OleDb
Public Class Menu_Utama
    Private Sub Menu_Utama_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Supir.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Petugas.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Pelanggan.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Mobil.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Peminjaman.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Pengembalian.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Laporan_Peminjaman.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Laporan_Pengembalian.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        GantiPassword.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If MessageBox.Show("Yakin Ingin Log Off?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            Login.Show()
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If MessageBox.Show("Yakin Ingin Keluar?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub SupirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupirToolStripMenuItem.Click
        Supir.Show()
    End Sub

    Private Sub PetugasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PetugasToolStripMenuItem.Click
        Petugas.Show()
    End Sub

    Private Sub PelangganToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PelangganToolStripMenuItem.Click
        Pelanggan.Show()
    End Sub

    Private Sub MobilToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MobilToolStripMenuItem.Click
        Mobil.Show()
    End Sub

    Private Sub PeminjamanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PeminjamanToolStripMenuItem.Click
        Peminjaman.Show()
    End Sub

    Private Sub PengembalianToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PengembalianToolStripMenuItem.Click
        Pengembalian.Show()
    End Sub

    Private Sub GantiPasswordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        GantiPassword.Show()
    End Sub

    Private Sub LogOffToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogOffToolStripMenuItem.Click
        Me.Hide()
        Login.Show()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        LaporanDataMaster.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        BackupDatabase.Show()
    End Sub

    Private Sub BackupDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BackupDataToolStripMenuItem.Click
        BackupDatabase.Show()
    End Sub
End Class