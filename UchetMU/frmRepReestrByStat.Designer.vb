<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRepReestrByStat
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.OK_Button = New System.Windows.Forms.Button
        Me.cbPerFrom = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbVidReport = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbPodr = New System.Windows.Forms.ComboBox
        Me.AllPodr = New System.Windows.Forms.RadioButton
        Me.optPodr = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbPerTo = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(340, 219)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(195, 36)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(101, 4)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(89, 28)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Отмена"
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK_Button.Location = New System.Drawing.Point(4, 4)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(89, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'cbPerFrom
        '
        Me.cbPerFrom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPerFrom.FormattingEnabled = True
        Me.cbPerFrom.Location = New System.Drawing.Point(101, 64)
        Me.cbPerFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbPerFrom.Name = "cbPerFrom"
        Me.cbPerFrom.Size = New System.Drawing.Size(189, 24)
        Me.cbPerFrom.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 64)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Период c"
        '
        'cbVidReport
        '
        Me.cbVidReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbVidReport.FormattingEnabled = True
        Me.cbVidReport.Items.AddRange(New Object() {"наличный расчет", "безналичный расчет", "военнослужащие", "ДМС", "УМО", "Все"})
        Me.cbVidReport.Location = New System.Drawing.Point(161, 18)
        Me.cbVidReport.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbVidReport.Name = "cbVidReport"
        Me.cbVidReport.Size = New System.Drawing.Size(225, 24)
        Me.cbVidReport.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 18)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Источник средств:"
        '
        'cbPodr
        '
        Me.cbPodr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPodr.Enabled = False
        Me.cbPodr.FormattingEnabled = True
        Me.cbPodr.Location = New System.Drawing.Point(149, 48)
        Me.cbPodr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbPodr.Name = "cbPodr"
        Me.cbPodr.Size = New System.Drawing.Size(315, 24)
        Me.cbPodr.TabIndex = 6
        '
        'AllPodr
        '
        Me.AllPodr.AutoSize = True
        Me.AllPodr.Checked = True
        Me.AllPodr.Location = New System.Drawing.Point(12, 21)
        Me.AllPodr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.AllPodr.Name = "AllPodr"
        Me.AllPodr.Size = New System.Drawing.Size(50, 21)
        Me.AllPodr.TabIndex = 8
        Me.AllPodr.TabStop = True
        Me.AllPodr.Text = "Все"
        Me.AllPodr.UseVisualStyleBackColor = True
        '
        'optPodr
        '
        Me.optPodr.AutoSize = True
        Me.optPodr.Location = New System.Drawing.Point(12, 48)
        Me.optPodr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.optPodr.Name = "optPodr"
        Me.optPodr.Size = New System.Drawing.Size(131, 21)
        Me.optPodr.TabIndex = 9
        Me.optPodr.Text = "Подразделение"
        Me.optPodr.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbPodr)
        Me.GroupBox1.Controls.Add(Me.optPodr)
        Me.GroupBox1.Controls.Add(Me.AllPodr)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 107)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(477, 90)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Подразделения"
        '
        'cbPerTo
        '
        Me.cbPerTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPerTo.FormattingEnabled = True
        Me.cbPerTo.Location = New System.Drawing.Point(343, 64)
        Me.cbPerTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbPerTo.Name = "cbPerTo"
        Me.cbPerTo.Size = New System.Drawing.Size(189, 24)
        Me.cbPerTo.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(309, 68)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 17)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "по"
        '
        'frmRepReestrByStat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(548, 268)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cbPerTo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbVidReport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbPerFrom)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRepReestrByStat"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Отчет: услуги по статьям калькуляции"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents cbPerFrom As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbVidReport As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbPodr As System.Windows.Forms.ComboBox
    Friend WithEvents AllPodr As System.Windows.Forms.RadioButton
    Friend WithEvents optPodr As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbPerTo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
