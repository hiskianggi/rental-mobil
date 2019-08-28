Imports System.Data.OleDb
Public Class Supir
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from DataSupir", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Sub Clear_Text()
        Input1.Clear()
        Input2.Clear()
        Input3.Clear()
        Input4.Clear()
        Input5.Clear()
        Input6.Clear()
        Input7.Clear()
        Input8.Clear()
        Input9.Text = "Tersedia"
        TextBox1.Clear()
        TampilGrid()
    End Sub
    Private Sub FormSupir_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilGrid()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input4.Text = "" Or Input3.Text = "" Or Input5.Text = "" Or Input6.Text = "" Or Input7.Text = "" Or Input8.Text = "" Or Input9.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Salah Satu atau Beberapa Kolom Belum Terisi!!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            Call koneksi()
            Dim simpan As String
            simpan = "insert into DataSupir values('" & Input1.Text & "','" & Input2.Text & "','" & Input3.Text & "','" & Input4.Text & "','" & Input5.Text & "','" & Input6.Text & "','" & Input7.Text & "','" & Input8.Text & "','" & Input9.Text & "','" & TextBox1.Text & "')"
            cmd = New OleDbCommand(simpan, conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Clear_Text()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Clear_Text()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        Input1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Input2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Input3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Input4.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Input5.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        Input6.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        Input7.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        Input8.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        Input9.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(9).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Input1.Text = "" Then
            MessageBox.Show("Pilih Salah Satu Data Untuk Dihapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "         ID = " & Input1.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from DataSupir where IDSupir ='" & Input1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Clear_Text()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub Input3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input3.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled() = True
    End Sub
    Private Sub Input6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input6.KeyPress
        If Not ((((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = " ") Or e.KeyChar = vbBack)) Or ((e.KeyChar = "(") Or e.KeyChar = ")"))) Then e.Handled() = True
    End Sub
    Private Sub Input7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input7.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = "+") Or e.KeyChar = vbBack)) Then e.Handled() = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input3.Text = "" Or Input4.Text = "" Or Input5.Text = "" Or Input6.Text = "" Or Input7.Text = "" Or Input8.Text = "" Or Input9.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Salah Satu Data Belum Terisi!!")
        Else
            Call koneksi()
            cmd = New OleDbCommand("update DataSupir SET IDSupir='" & Input1.Text & "',Nama='" & Input2.Text & "',NoSIM='" & Input3.Text & "',JenisKelamin='" & Input4.Text & "',Alamat='" & Input5.Text & "',TelpRumah='" & Input6.Text & "',Handphone='" & Input7.Text & "',Email='" & Input8.Text & "',Status='" & Input9.Text & "',HargaSewa='" & TextBox1.Text & "' WHERE IDSupir='" & Input1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Clear_Text()
            TampilGrid()
        End If
    End Sub
End Class