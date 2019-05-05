Imports System.Data.OLEDB
Imports System.Windows.Forms

Public Class frmFilterPeriod
    Public Per As New modBas.Period
    Private adaptercombo As OleDbDataAdapter
    Private dscombo As New DataTable

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmFilterPeriod_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from D_Periods", Connection)
        adaptercombo.Fill(dscombo)
        cbPeriods.DataSource = dscombo
        cbPeriods.ValueMember = "ID"
        cbPeriods.DisplayMember = "Наименование"
    End Sub
End Class
