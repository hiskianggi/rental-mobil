<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BackupDatabase
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.FileListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.DirListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.DriveListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DirListBox2 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.DriveListBox2 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.FileListBox1)
        Me.GroupBox1.Controls.Add(Me.DirListBox1)
        Me.GroupBox1.Controls.Add(Me.DriveListBox1)
        Me.GroupBox1.Font = New System.Drawing.Font("Baar Sophia", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(25, 93)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(477, 306)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "File Sumber"
        '
        'FileListBox1
        '
        Me.FileListBox1.FormattingEnabled = True
        Me.FileListBox1.Location = New System.Drawing.Point(303, 37)
        Me.FileListBox1.Name = "FileListBox1"
        Me.FileListBox1.Pattern = "*.*"
        Me.FileListBox1.Size = New System.Drawing.Size(155, 257)
        Me.FileListBox1.TabIndex = 2
        '
        'DirListBox1
        '
        Me.DirListBox1.FormattingEnabled = True
        Me.DirListBox1.IntegralHeight = False
        Me.DirListBox1.Location = New System.Drawing.Point(22, 77)
        Me.DirListBox1.Name = "DirListBox1"
        Me.DirListBox1.Size = New System.Drawing.Size(264, 214)
        Me.DirListBox1.TabIndex = 1
        '
        'DriveListBox1
        '
        Me.DriveListBox1.FormattingEnabled = True
        Me.DriveListBox1.Location = New System.Drawing.Point(23, 39)
        Me.DriveListBox1.Name = "DriveListBox1"
        Me.DriveListBox1.Size = New System.Drawing.Size(263, 32)
        Me.DriveListBox1.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DirListBox2)
        Me.GroupBox2.Controls.Add(Me.DriveListBox2)
        Me.GroupBox2.Font = New System.Drawing.Font("Baar Sophia", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(523, 93)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(325, 306)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Direktori Tujuan"
        '
        'DirListBox2
        '
        Me.DirListBox2.FormattingEnabled = True
        Me.DirListBox2.IntegralHeight = False
        Me.DirListBox2.Location = New System.Drawing.Point(33, 77)
        Me.DirListBox2.Name = "DirListBox2"
        Me.DirListBox2.Size = New System.Drawing.Size(264, 214)
        Me.DirListBox2.TabIndex = 2
        '
        'DriveListBox2
        '
        Me.DriveListBox2.FormattingEnabled = True
        Me.DriveListBox2.Location = New System.Drawing.Point(33, 39)
        Me.DriveListBox2.Name = "DriveListBox2"
        Me.DriveListBox2.Size = New System.Drawing.Size(263, 32)
        Me.DriveListBox2.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(238, 423)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(401, 29)
        Me.TextBox1.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(238, 469)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(401, 29)
        Me.TextBox2.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Highlight
        Me.Button1.Font = New System.Drawing.Font("Baar Sophia", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button1.Location = New System.Drawing.Point(682, 423)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(137, 75)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Backup"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Baar Sophia", 15.75!)
        Me.Label1.Location = New System.Drawing.Point(55, 429)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "File Sumber"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Baar Sophia", 15.75!)
        Me.Label2.Location = New System.Drawing.Point(55, 472)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(147, 23)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Direktori Tujuan"
        '
        'BackupDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.BackgroundImage = Global.Hiskia___Rental_Mobil.My.Resources.Resources.R1
        Me.ClientSize = New System.Drawing.Size(875, 516)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "BackupDatabase"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Backup Database"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents FileListBox1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents DirListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DirListBox2 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox2 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
