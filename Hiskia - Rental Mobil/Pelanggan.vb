Imports System.Data.OleDb
Public Class Pelanggan
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from Pelanggan", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Sub BersihkanTeks()
        Input1.Clear()
        Input2.Clear()
        Input3.Clear()
        Input5.Clear()
        Input6.Clear()
        Input7.Clear()
        Input8.Clear()
        Input9.Clear()
        Input10.Clear()
        Input11.Clear()
        Input12.Clear()
        Input13.Clear()
    End Sub
    Private Sub FormPelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilGrid()
        BersihkanTeks()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input3.Text = "" Or Input5.Text = "" Or Input6.Text = "" Or Input7.Text = "" Or Input8.Text = "" Or Input9.Text = "" Or Input10.Text = "" Or Input11.Text = "" Or Input12.Text = "" Or Input13.Text = "" Then
            MessageBox.Show("Salah Satu atau Beberapa Kolom Belum Terisi!!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If RadioButton2.Checked Then
                Call koneksi()
                Dim simpan As String
                simpan = "insert into Pelanggan values('" & Input1.Text & "','" & Input2.Text & "','" & Input3.Text & "','" & RadioButton2.Text & "','" & Input5.Text & "','" & Input6.Text & "','" & Input7.Text & "','" & Input8.Text & "','" & Input9.Text & "','" & Input10.Text & "','" & Input11.Text & "','" & Input12.Text & "','" & Input13.Text & "')"
                cmd = New OleDbCommand(simpan, conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                BersihkanTeks()
                TampilGrid()
            ElseIf RadioButton1.Checked Then
                Call koneksi()
                Dim simpan As String
                simpan = "insert into Pelanggan values('" & Input1.Text & "','" & Input2.Text & "','" & Input3.Text & "','" & RadioButton1.Text & "','" & Input5.Text & "','" & Input6.Text & "','" & Input7.Text & "','" & Input8.Text & "','" & Input9.Text & "','" & Input10.Text & "','" & Input11.Text & "','" & Input12.Text & "','" & Input13.Text & "')"
                cmd = New OleDbCommand(simpan, conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                BersihkanTeks()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        BersihkanTeks()
        TampilGrid()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        Input1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Input2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Input3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Input5.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        Input6.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        Input7.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        Input8.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        Input9.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value
        Input10.Text = DataGridView1.Rows(e.RowIndex).Cells(9).Value
        Input11.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value
        Input12.Text = DataGridView1.Rows(e.RowIndex).Cells(11).Value
        Input13.Text = DataGridView1.Rows(e.RowIndex).Cells(12).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Input1.Text = "" Then
            MessageBox.Show("Anda Belum Memilih Data Yang Akan Dihapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "             ID Pelanggan = " & Input1.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from Pelanggan where IDPelanggan='" & Input1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Information)
                BersihkanTeks()
                TampilGrid()
            End If
        End If
    End Sub

    Private Sub Input3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input3.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled() = True
    End Sub

    Private Sub Input10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input10.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled() = True
    End Sub

    Private Sub Input11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input11.KeyPress
        If Not ((((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = " ") Or e.KeyChar = vbBack)) Or ((e.KeyChar = "(") Or e.KeyChar = ")"))) Then e.Handled() = True
    End Sub

    Private Sub Input12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Input12.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = "+") Or e.KeyChar = vbBack)) Then e.Handled() = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Input1.Text = "" Or Input2.Text = "" Or Input3.Text = "" Or Input5.Text = "" Or Input6.Text = "" Or Input7.Text = "" Or Input8.Text = "" Or Input9.Text = "" Or Input10.Text = "" Or Input11.Text = "" Or Input12.Text = "" Or Input13.Text = "" Then
            MessageBox.Show("Salah Satu Data Belum Terisi!!")
        ElseIf RadioButton1.Checked Then
            Call koneksi()
            cmd = New OleDbCommand("update Pelanggan SET IDPelanggan='" & Input1.Text & "',Nama='" & Input2.Text & "',NoKTP='" & Input3.Text & "',JenisKelamin='" & RadioButton1.Text & "',TempatLahir='" & Input5.Text & "',TanggalLahir='" & Input6.Text & "',Pekerjaan='" & Input7.Text & "',Alamat='" & Input8.Text & "',Kota='" & Input9.Text & "',KodePos='" & Input10.Text & "',TelpRumah='" & Input11.Text & "',Handphone='" & Input12.Text & "',Email='" & Input13.Text & "' WHERE IDPelanggan='" & Input1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        ElseIf RadioButton2.Checked Then
            Call koneksi()
            cmd = New OleDbCommand("update Pelanggan SET IDPelanggan='" & Input1.Text & "',Nama='" & Input2.Text & "',NoKTP='" & Input3.Text & "',JenisKelamin='" & RadioButton2.Text & "',TempatLahir='" & Input5.Text & "',TanggalLahir='" & Input6.Text & "',Pekerjaan='" & Input7.Text & "',Alamat='" & Input8.Text & "',Kota='" & Input9.Text & "',KodePos='" & Input10.Text & "',TelpRumah='" & Input11.Text & "',Handphone='" & Input12.Text & "',Email='" & Input13.Text & "' WHERE IDPelanggan='" & Input1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            BersihkanTeks()
            TampilGrid()
        End If
    End Sub
End Class