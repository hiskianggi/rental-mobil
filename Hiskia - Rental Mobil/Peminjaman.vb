Imports System.Data.OleDb
Public Class Peminjaman
    Dim proses As New ClsKoneksi
    Dim dttble As DataTable
    Sub data()
        iddantanggal()
        koneksi()
        TampilGrid()
        Call koneksi()
        Qry = "select * from Pelanggan"
        cmd = New OleDbCommand(Qry, conn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item("IDPelanggan"))
        Loop
        Call koneksi()
        MobilQry = "select * from DataMobil"
        cmd = New OleDbCommand(MobilQry, conn)
        dr = cmd.ExecuteReader
        ComboBox2.Items.Clear()
        Do While dr.Read
            ComboBox2.Items.Add(dr.Item("IDMobil"))
        Loop
        Call koneksi()
        SupirQry = "select * from DataSupir"
        cmd = New OleDbCommand(SupirQry, conn)
        dr = cmd.ExecuteReader
        ComboBox4.Items.Clear()
        Do While dr.Read
            ComboBox4.Items.Add(dr.Item("IDSupir"))
        Loop
    End Sub
    Sub Tanggal()
        Dim tanggal As Date
        Dim hasil As String
        tanggal = Date.FromOADate(Format(Now, "dd") + Val(TextBox4.Text))
        hasil = tanggal
        TextBox6.Text = hasil
    End Sub
    Sub iddantanggal()
        Dim hasil As Integer
        Try
            dttble = proses.ExecuteQuery("Select * From Pinjam order by NomorPinjam desc")
            If dttble.Rows.Count = 0 Then
                TextBox10.Text = "" + Format(Now, "yyddMM") + "001"
            Else
                With dttble.Rows(0)
                    TextBox10.Text = .Item("NomorPinjam")
                End With

                hasil = Val(Microsoft.VisualBasic.Mid(TextBox10.Text, 12, 3)) + Val(1)
                TextBox10.Text = hasil
                If Len(TextBox10.Text) = 1 Then
                    TextBox10.Text = "" + Format(Now, "yyddMM") + "00" & TextBox10.Text & ""
                ElseIf Len(TextBox10.Text) = 2 Then
                    TextBox10.Text = "" + Format(Now, "yyddMM") + "0" & TextBox10.Text & ""
                ElseIf Len(TextBox10.Text) = 3 Then
                    TextBox10.Text = "" + Format(Now, "yyddMM") + "" & TextBox10.Text & ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub TampilGrid()
        Call koneksi()
        da = New OleDbDataAdapter("select * from DataMobil", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox15.Text = Today
        TextBox16.Text = TimeOfDay
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        TextBox15.Text = Today
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        TextBox16.Text = TimeOfDay
    End Sub

    Private Sub FormPeminjaman_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        data()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call koneksi()
        Qry = "select * from Pelanggan where IDPelanggan='" & ComboBox1.Text & "'"
        cmd = New OleDbCommand(Qry, conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Input1.Text = dr.Item("Nama")
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Call koneksi()
        SupirQry = "select * from DataSupir where IDSupir='" & ComboBox4.Text & "'"
        cmd = New OleDbCommand(SupirQry, conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox9.Text = dr.Item("Nama")
            TextBox14.Text = dr.Item("HargaSewa")
        End If


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call koneksi()
        MobilQry = "select * from DataMobil where IDMobil='" & ComboBox2.Text & "'"
        cmd = New OleDbCommand(MobilQry, conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox2.Text = dr.Item("MerkMobil")
            TextBox1.Text = dr.Item("NoPolisi")
            TextBox3.Text = dr.Item("HargaSewa")
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
    End Sub
    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If Not TextBox3.Text = "" And TextBox4.Text = "" And TextBox14.Text = "" Then
            TextBox7.Text = "-"
        Else
            Dim Hasil As Integer
            Hasil = Val(TextBox3.Text) + Val(TextBox14.Text)
            TextBox7.Text = Val(Hasil) * Val(TextBox4.Text)
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F7 Then
            Dim sisabayar As Integer
            sisabayar = Val(TextBox7.Text) - Val(TextBox8.Text)
            TextBox5.Text = sisabayar
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        TextBox13.Text = TextBox8.Text
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F7 Then
            Dim kembalian As Integer
            kembalian = Val(TextBox11.Text) - Val(TextBox13.Text)
            TextBox12.Text = kembalian
        End If
    End Sub
    Sub clear()
        data()
        Input1.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox9.Clear()
        TextBox13.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox14.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        clear()
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Tanggal()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        Dim simpan As String
        simpan = "insert into Pinjam values('" & TextBox10.Text & "','" & TextBox15.Text & "','" & ComboBox1.Text & "','" & Input1.Text & "','" & ComboBox2.Text & "','" & TextBox2.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox4.Text & "','" & TextBox9.Text & "','" & TextBox14.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox13.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "')"
        cmd = New OleDbCommand(simpan, conn)
        cmd.ExecuteNonQuery()
        MessageBox.Show("Tersimpan!", "Hiskia - Rental Mobil", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        clear()
    End Sub
End Class