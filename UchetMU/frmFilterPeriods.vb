Imports System.Data.OLEDB
Imports System.Windows.Forms

Public Class frmFilterPeriods
    Private adaptercombo As OleDbDataAdapter
    Private dscombo As New DataTable
    Private dscomboTo As New DataTable
    Public PerFrom As Period
    Public PerTo As Period

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        PerFrom.ID = cbPerFrom.SelectedValue
        PerTo.ID = cbPerTo.SelectedValue
        If PerFrom.dNach > PerTo.dNach Then
            MsgBox("�������� ����� �����������.")
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmFilterPeriods_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from D_Periods", Connection)
        adaptercombo.Fill(dscombo)
        cbPerFrom.DataSource = dscombo
        cbPerFrom.ValueMember = "ID"
        cbPerFrom.DisplayMember = "������������"
        adaptercombo.Fill(dscomboTo)
        cbPerTo.DataSource = dscomboTo
        cbPerTo.ValueMember = "ID"
        cbPerTo.DisplayMember = "������������"
    End Sub
End Class
