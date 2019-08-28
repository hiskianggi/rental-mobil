Imports System.Data.OleDb
Public Class Mobil
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from DataMobil", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub

    Sub Clear_Text()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        ComboBox7.Text = ""
        Input8.Clear()
        TampilGrid()
        TextBox1.Focus()
    End Sub
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        TampilGrid()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox7.Text = "" Then
            MessageBox.Show("Salah Satu Data Belum Terisi!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            Call koneksi()
            Dim simpan As String
            simpan = "insert into DataMobil values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & ComboBox7.Text & "')"
            cmd = New OleDbCommand(simpan, conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Clear_Text()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Clear_Text()
        TampilGrid()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
        Clear_Text()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" And TextBox6.Text = "" And ComboBox7.Text = "" Then
            MessageBox.Show("Pilih Salah Satu Data Yang Akan Dihapus!!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            If MessageBox.Show("Yakin Ingin Menghapus" & vbNewLine & "" & vbNewLine & "         IDMobil = " & TextBox2.Text & "?", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("delete from DataMobil where IDMobil='" & TextBox1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Terhapus!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Clear_Text()
                TampilGrid()
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        ComboBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
    End Sub

    Private Sub Input8_KeyDown(sender As Object, e As KeyEventArgs) Handles Input8.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Input8.Text = "" Then
                MessageBox.Show("Masukkan Isi Pencarian!!")
            Else
                cmd = New OleDbCommand("select * from DataMobil where " & Label1.Text & " like '%" & Input8.Text & "%'", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                koneksi()
                da = New OleDbDataAdapter("select * from DataMobil where " & Label1.Text & " like '%" & Input8.Text & "%'", conn)
                ds = New DataSet
                da.Fill(ds)
                If dr.HasRows Then
                    ds = New DataSet
                    da.Fill(ds)
                    DataGridView1.DataSource = ds.Tables(0)
                Else
                    MessageBox.Show("Item Tidak Ditemukan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Select()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or ((e.KeyChar = "+") Or e.KeyChar = vbBack)) Then e.Handled() = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or ComboBox7.Text = "" Then
            MessageBox.Show("Salah Satu Data Belum Terisi!!")
        Else
            Call koneksi()
            cmd = New OleDbCommand("update DataMobil SET IDMobil='" & TextBox1.Text & "',MerkMobil='" & TextBox2.Text & "',Warna='" & TextBox3.Text & "',Tahun='" & TextBox4.Text & "',NoPolisi='" & TextBox5.Text & "',HargaSewa='" & TextBox6.Text & "',Status='" & ComboBox7.Text & "' WHERE IDMobil='" & TextBox1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Updated Sucess!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Clear_Text()
            TampilGrid()
        End If
    End Sub
End Class