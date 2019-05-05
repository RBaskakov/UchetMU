<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAbout
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub


    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LabelCopyRight = New System.Windows.Forms.Label()
        Me.butOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.SuspendLayout()
        '
        'LabelCopyRight
        '
        Me.LabelCopyRight.AutoSize = True
        Me.LabelCopyRight.Location = New System.Drawing.Point(12, 80)
        Me.LabelCopyRight.Name = "LabelCopyRight"
        Me.LabelCopyRight.Size = New System.Drawing.Size(39, 13)
        Me.LabelCopyRight.TabIndex = 3
        Me.LabelCopyRight.Text = "Label1"
        '
        'butOK
        '
        Me.butOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butOK.Location = New System.Drawing.Point(52, 117)
        Me.butOK.Name = "butOK"
        Me.butOK.Size = New System.Drawing.Size(87, 33)
        Me.butOK.TabIndex = 4
        Me.butOK.Text = "OK"
        Me.butOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(11, 19)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 29)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "По вопросам приобретения и поддержки  обращайтесь:"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(44, 49)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(100, 13)
        Me.LinkLabel1.TabIndex = 6
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "roman_box@mail.ru"
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(190, 169)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.butOK)
        Me.Controls.Add(Me.LabelCopyRight)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "О программе"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelCopyRight As System.Windows.Forms.Label
    Friend WithEvents butOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel

End Class
