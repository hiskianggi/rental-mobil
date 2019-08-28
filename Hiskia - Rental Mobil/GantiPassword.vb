Imports System.Data.OleDb
Public Class GantiPassword
    Private Sub GantiPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        bersihkanteks()
        TextBox2.PasswordChar = "*"
        TextBox3.PasswordChar = "*"
        TextBox4.PasswordChar = "*"
        TextBox1.Focus()
    End Sub
    Sub bersihkanteks()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text <> TextBox4.Text Then
            MessageBox.Show("Password Konfirmasi Tidak Sama!!")
            TextBox4.Clear()
            TextBox4.Focus()
        Else
            If MessageBox.Show("Yakin Ingin Mengganti Password Anda? Klik 'Yes' Untuk Melanjutkan...", "Peringatan!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                cmd = New OleDbCommand("update Petugas set KataSandi='" & TextBox4.Text & "'where Nama='" & TextBox1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Updated Sucess!", "Hiskia Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                bersihkanteks()
                TextBox1.Focus()
            Else
                bersihkanteks()
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        bersihkanteks()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            Call koneksi()
            If conn.State <> ConnectionState.Closed Then conn.Close()
            conn.Open()
            cmd = New OleDbCommand("SELECT * From Petugas  WHERE Nama = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            If Not dr.HasRows Then
                MessageBox.Show("Nama User Salah!!")
                TextBox1.Clear()
                TextBox1.Focus()
            Else
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
            Call koneksi()
            If conn.State <> ConnectionState.Closed Then conn.Close()
            conn.Open()
            cmd = New OleDbCommand("SELECT * From Petugas  WHERE Nama = '" & TextBox1.Text & "' and KataSandi = '" & TextBox2.Text & "'", conn)
            dr = cmd.ExecuteReader
            If Not dr.HasRows Then
                MessageBox.Show("Password Salah!!")
                TextBox2.Clear()
                TextBox2.Focus()
            Else
                TextBox3.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
            If TextBox3.Text = "" Then
                MessageBox.Show("Isi Password Baru Anda!!")
                TextBox3.Clear()
                TextBox3.Focus()
            Else
                TextBox4.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub

End Class