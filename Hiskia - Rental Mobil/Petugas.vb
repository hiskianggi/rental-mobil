Imports System.Data.OleDb
Public Class Petugas
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from Petugas", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub FormPetugas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilGrid()
    End Sub

    Sub BersihkanTeks()
        Input1.Clear()
        Input2.Clear()
        Input3.Clear()
        Input4.Clear()
        TampilGrid()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input4.Text = "" Or Input3.Text = "" Then
            MessageBox.Show("Salah Satu atau Beberapa Kolom Belum Terisi!!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            Call koneksi()
            Dim simpan As String
            simpan = "insert into Petugas values('" & Input1.Text & "','" & Input2.Text & "','" & Input3.Text & "','" & Input4.Text & "')"
            cmd = New OleDbCommand(simpan, conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        BersihkanTeks()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Input1.Text = "" Then
            MessageBox.Show("Pilih Salah Satu Data Untuk Dihapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "         ID = " & Input1.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from Petugas where ID='" & Input1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Information)
                BersihkanTeks()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        Input1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Input2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Input3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Input4.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input4.Text = "" Then
            MessageBox.Show("")
        Else
            Call koneksi()
            cmd = New OleDbCommand("update Petugas SET ID='" & Input1.Text & "',Nama='" & Input2.Text & "',Status='" & Input4.Text & "' WHERE ID='" & Input1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub
End Class