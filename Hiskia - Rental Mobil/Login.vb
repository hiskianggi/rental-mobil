Imports System.Data.OleDb
Public Class Login
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        If conn.State <> ConnectionState.Closed Then conn.Close()
        conn.Open()
        cmd = New OleDbCommand("SELECT * From Petugas  WHERE Nama = '" & TextBox1.Text & "' and KataSandi = '" & TextBox2.Text & "'", conn)
        dr = cmd.ExecuteReader()
        If (dr.HasRows()) Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
            Me.Hide()
            Mengalihkan.Show()
        Else
            MsgBox("Username & Password Anda Salah!", MessageBoxButtons.OK, "Login gagal")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        koneksi()
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call koneksi()
            If conn.State <> ConnectionState.Closed Then conn.Close()
            conn.Open()
            cmd = New OleDbCommand("SELECT * From Petugas  WHERE Nama = '" & TextBox1.Text & "' and KataSandi = '" & TextBox2.Text & "'", conn)
            dr = cmd.ExecuteReader()
            If (dr.HasRows()) Then
                Me.Hide()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
                Mengalihkan.Show()
            Else
                MsgBox("Username & Password Anda Salah!", MessageBoxButtons.OK, "Login gagal")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        TextBox1.Focus()
    End Sub
End Class