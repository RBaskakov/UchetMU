<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDB
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents txtPathDB As System.Windows.Forms.TextBox
    Friend WithEvents OK As System.Windows.Forms.Button
    Friend WithEvents Cancel As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.UsernameLabel = New System.Windows.Forms.Label
        Me.txtPathDB = New System.Windows.Forms.TextBox
        Me.OK = New System.Windows.Forms.Button
        Me.Cancel = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'UsernameLabel
        '
        Me.UsernameLabel.Location = New System.Drawing.Point(9, 11)
        Me.UsernameLabel.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(106, 19)
        Me.UsernameLabel.TabIndex = 0
        Me.UsernameLabel.Text = "Файл базы данных:"
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPathDB
        '
        Me.txtPathDB.Location = New System.Drawing.Point(110, 11)
        Me.txtPathDB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPathDB.Name = "txtPathDB"
        Me.txtPathDB.Size = New System.Drawing.Size(366, 20)
        Me.txtPathDB.TabIndex = 1
        '
        'OK
        '
        Me.OK.Location = New System.Drawing.Point(310, 46)
        Me.OK.Margin = New System.Windows.Forms.Padding(2)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(70, 22)
        Me.OK.TabIndex = 4
        Me.OK.Text = "&Сохранить"
        '
        'Cancel
        '
        Me.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel.Location = New System.Drawing.Point(404, 46)
        Me.Cancel.Margin = New System.Windows.Forms.Padding(2)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(70, 22)
        Me.Cancel.TabIndex = 5
        Me.Cancel.Text = "&Отмена"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(479, 10)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(27, 19)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "*"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmDB
        '
        Me.AcceptButton = Me.OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel
        Me.ClientSize = New System.Drawing.Size(515, 85)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.txtPathDB)
        Me.Controls.Add(Me.UsernameLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDB"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Выберите файл базы данных"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
