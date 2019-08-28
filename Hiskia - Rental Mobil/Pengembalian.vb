Imports System.Data.OleDb
Public Class Pengembalian
    Dim proses As New ClsKoneksi
    Dim dttble As DataTable
    Sub iddantanggal()
        Try
            dttble = proses.ExecuteQuery("Select * From Kembali order by NomorKembali desc")
            If dttble.Rows.Count = 0 Then
                TextBox2.Text = "" + Format(Now, "yyddMM") + "001"
            Else
                With dttble.Rows(0)
                    TextBox2.Text = .Item("NomorKembali")
                End With

                TextBox2.Text = Val(Microsoft.VisualBasic.Mid(TextBox2.Text, 12, 3)) + 1
                If Len(TextBox2.Text) = 1 Then
                    TextBox2.Text = "" + Format(Now, "dd/MM/yyyy") + "-00" & TextBox2.Text & ""
                ElseIf Len(TextBox2.Text) = 2 Then
                    TextBox2.Text = "" + Format(Now, "dd/MM/yyyy") + "-0" & TextBox2.Text & ""
                ElseIf Len(TextBox2.Text) = 3 Then
                    TextBox2.Text = "" + Format(Now, "dd/MM/yyyy") + "" & TextBox2.Text & ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Pengembalian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        iddantanggal()
    End Sub
    Sub tampilgrid2()
        Call koneksi()
        da = New OleDbDataAdapter("select * from Kembali", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView2.DataSource = ds.Tables(0)
        DataGridView2.ReadOnly = True
    End Sub
    Sub tampilgrid1()
        Call koneksi()
        da = New OleDbDataAdapter("select * from Pinjam", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub FormPengembalian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        tampilgrid1()
        tampilgrid2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox3.Text = Today
        TextBox4.Text = TimeOfDay
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.Text = Today
        TextBox4.Text = TimeOfDay
    End Sub
End Class