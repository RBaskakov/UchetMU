<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReport
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(91, 22)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(93, 33)
        Me.ButtonCancel.TabIndex = 0
        Me.ButtonCancel.Text = "Отмена"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'frmReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(265, 89)
        Me.Controls.Add(Me.ButtonCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReport"
        Me.RightToLeftLayout = True
        Me.Text = "Формирование отчёта"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ButtonCancel As Windows.Forms.Button
End Class
