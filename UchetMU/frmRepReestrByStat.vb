Imports System.Data.OLEDB
Imports System.Windows.Forms

Public Class frmRepReestrByStat
    Private adaptercombo As OleDbDataAdapter
    Private dscombo, dscomboTo As New DataTable
    Private adaptercomboPodr As OleDbDataAdapter
    Private dscomboPodr As New DataTable
    Public PerFrom As Period
    Public PerTo As Period

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        PerFrom.ID = cbPerFrom.SelectedValue
        PerTo.ID = cbPerTo.SelectedValue
        If PerFrom.dNach > PerTo.dNach Then
            MsgBox("Интервал задан некорректно.")
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmRepReestrByStat_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dRow As DataRow
        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from D_Periods", Connection)
        adaptercombo.Fill(dscombo)
        cbPerFrom.DataSource = dscombo
        cbPerFrom.ValueMember = "ID"
        cbPerFrom.DisplayMember = "Наименование"
        adaptercombo.Fill(dscomboTo)
        cbPerTo.DataSource = dscomboTo
        cbPerTo.ValueMember = "ID"
        cbPerTo.DisplayMember = "Наименование"

        dscomboPodr.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercomboPodr = New OleDbDataAdapter("select * from D_Podr order by Подразделение", Connection)
        adaptercomboPodr.Fill(dscomboPodr)
        dRow = dscomboPodr.NewRow
        dRow("ID") = -1
        dRow("Код_подразделения") = 0
        dRow("Подразделение") = "Амбулаторно-поликлиническая помощь"
        dscomboPodr.Rows.Add(dRow)
        cbPodr.DataSource = dscomboPodr
        cbPodr.ValueMember = "ID"
        cbPodr.DisplayMember = "Подразделение"

    End Sub

    Private Sub optPodr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPodr.CheckedChanged
        cbPodr.Enabled = optPodr.Checked
    End Sub

    Private Sub cbPodr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPodr.SelectedIndexChanged
        If cbVidReport.SelectedIndex > -1 And cbVidReport.SelectedIndex <> 2 Then
            If cbPodr.SelectedValue = -1 Then cbPodr.SelectedIndex = 0
        End If
    End Sub

    Private Sub cbVidReport_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVidReport.SelectedIndexChanged
        If cbVidReport.SelectedIndex > -1 And cbVidReport.SelectedIndex <> 2 Then
            If cbPodr.SelectedValue = -1 Then cbPodr.SelectedIndex = 0
        End If
    End Sub

    
End Class
