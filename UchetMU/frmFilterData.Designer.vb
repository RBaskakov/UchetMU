<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFilterData
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
        Me.dBegin = New System.Windows.Forms.DateTimePicker
        Me.dEnd = New System.Windows.Forms.DateTimePicker
        Me.lblEnd = New System.Windows.Forms.Label
        Me.cbPeriods = New System.Windows.Forms.ComboBox
        Me.optInterval = New System.Windows.Forms.RadioButton
        Me.optPeriod = New System.Windows.Forms.RadioButton
        Me.optNoPay = New System.Windows.Forms.RadioButton
        Me.optAll = New System.Windows.Forms.RadioButton
        Me.TableLayoutPanel1.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(199, 173)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(251, 36)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(137, 4)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(101, 28)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Отмена"
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(5, 4)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(115, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "&Отчет"
        '
        'dBegin
        '
        Me.dBegin.Enabled = False
        Me.dBegin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dBegin.Location = New System.Drawing.Point(76, 62)
        Me.dBegin.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dBegin.Name = "dBegin"
        Me.dBegin.Size = New System.Drawing.Size(151, 22)
        Me.dBegin.TabIndex = 4
        Me.dBegin.Value = New Date(2008, 12, 18, 0, 0, 0, 0)
        '
        'dEnd
        '
        Me.dEnd.Enabled = False
        Me.dEnd.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dEnd.Location = New System.Drawing.Point(267, 63)
        Me.dEnd.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dEnd.Name = "dEnd"
        Me.dEnd.Size = New System.Drawing.Size(151, 22)
        Me.dEnd.TabIndex = 6
        '
        'lblEnd
        '
        Me.lblEnd.AutoSize = True
        Me.lblEnd.Location = New System.Drawing.Point(235, 68)
        Me.lblEnd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEnd.Name = "lblEnd"
        Me.lblEnd.Size = New System.Drawing.Size(24, 17)
        Me.lblEnd.TabIndex = 8
        Me.lblEnd.Text = "по"
        '
        'cbPeriods
        '
        Me.cbPeriods.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPeriods.Enabled = False
        Me.cbPeriods.FormattingEnabled = True
        Me.cbPeriods.Location = New System.Drawing.Point(123, 21)
        Me.cbPeriods.Margin = New System.Windows.Forms.Padding(4)
        Me.cbPeriods.Name = "cbPeriods"
        Me.cbPeriods.Size = New System.Drawing.Size(189, 24)
        Me.cbPeriods.TabIndex = 9
        '
        'optInterval
        '
        Me.optInterval.AutoSize = True
        Me.optInterval.Location = New System.Drawing.Point(28, 64)
        Me.optInterval.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.optInterval.Name = "optInterval"
        Me.optInterval.Size = New System.Drawing.Size(33, 21)
        Me.optInterval.TabIndex = 10
        Me.optInterval.Text = "c"
        Me.optInterval.UseVisualStyleBackColor = True
        '
        'optPeriod
        '
        Me.optPeriod.AutoSize = True
        Me.optPeriod.Location = New System.Drawing.Point(28, 21)
        Me.optPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.optPeriod.Name = "optPeriod"
        Me.optPeriod.Size = New System.Drawing.Size(80, 21)
        Me.optPeriod.TabIndex = 11
        Me.optPeriod.Text = "Период:"
        Me.optPeriod.UseVisualStyleBackColor = True
        '
        'optNoPay
        '
        Me.optNoPay.AutoSize = True
        Me.optNoPay.Location = New System.Drawing.Point(28, 111)
        Me.optNoPay.Margin = New System.Windows.Forms.Padding(4)
        Me.optNoPay.Name = "optNoPay"
        Me.optNoPay.Size = New System.Drawing.Size(103, 21)
        Me.optNoPay.TabIndex = 12
        Me.optNoPay.TabStop = True
        Me.optNoPay.Text = "Без оплаты"
        Me.optNoPay.UseVisualStyleBackColor = True
        '
        'optAll
        '
        Me.optAll.AutoSize = True
        Me.optAll.Checked = True
        Me.optAll.Location = New System.Drawing.Point(28, 153)
        Me.optAll.Name = "optAll"
        Me.optAll.Size = New System.Drawing.Size(50, 21)
        Me.optAll.TabIndex = 13
        Me.optAll.TabStop = True
        Me.optAll.Text = "Все"
        Me.optAll.UseVisualStyleBackColor = True
        '
        'frmFilterData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(463, 222)
        Me.Controls.Add(Me.optAll)
        Me.Controls.Add(Me.optNoPay)
        Me.Controls.Add(Me.optPeriod)
        Me.Controls.Add(Me.optInterval)
        Me.Controls.Add(Me.cbPeriods)
        Me.Controls.Add(Me.lblEnd)
        Me.Controls.Add(Me.dEnd)
        Me.Controls.Add(Me.dBegin)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFilterData"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Выберите интервал оплат"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents dBegin As System.Windows.Forms.DateTimePicker
    Friend WithEvents dEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblEnd As System.Windows.Forms.Label
    Friend WithEvents cbPeriods As System.Windows.Forms.ComboBox
    Friend WithEvents optInterval As System.Windows.Forms.RadioButton
    Friend WithEvents optPeriod As System.Windows.Forms.RadioButton
    Friend WithEvents optNoPay As System.Windows.Forms.RadioButton
    Friend WithEvents optAll As System.Windows.Forms.RadioButton

End Class
