Public Class BackupDatabase

    Private Sub DriveListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DriveListBox1.SelectedIndexChanged
        DirListBox1.Path = DriveListBox1.Drive
    End Sub

    Private Sub DirListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DirListBox1.SelectedIndexChanged
        'FileListBox1.Pattern = "(*.mdb) |*.mdb"
        FileListBox1.Path = DirListBox1.Path
    End Sub

    Private Sub FileListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FileListBox1.SelectedIndexChanged
        TextBox1.Text = FileListBox1.Path & "\" & FileListBox1.FileName
    End Sub

    Private Sub DriveListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DriveListBox2.SelectedIndexChanged
        DirListBox2.Path = DriveListBox2.Drive
    End Sub

    Private Sub DirListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DirListBox2.SelectedIndexChanged
        TextBox2.Text = DirListBox2.Path & "\" & FileListBox1.FileName
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Then
                MsgBox("Anda belum memilih database yang akan dibackup")
                Exit Sub
            ElseIf TextBox2.Text = "" Then
                MsgBox("Anda tidak memilih direktori tujuan backup")
                Exit Sub
            End If
            My.Computer.FileSystem.CopyFile(TextBox1.Text, TextBox2.Text)
            MsgBox("Backup Database Completed!")
            TextBox1.Clear()
            TextBox2.Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub BackupDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox2.Clear()
        DriveListBox1.Drive = "C:\"
        DriveListBox2.Drive = "C:\"
        FileListBox1.FileName = "db.mdb"
    End Sub
End Class