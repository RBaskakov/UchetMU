Imports System.Windows.Forms

Public Class frmCombo
    Public DataSource As DataTable
    Public SelectedValue As Integer
    Public ColumnName, STable, ColumnTitle As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmCombo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        combo.DataSource = DataSource
        combo.ValueMember = "ID"
        combo.DisplayMember = ColumnName
        combo.Name = STable
        combo.SelectedValue = SelectedValue
        lblCap.Text = ColumnTitle
    End Sub

End Class
