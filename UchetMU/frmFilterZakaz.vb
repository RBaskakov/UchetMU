Imports System.Windows.Forms

Public Class frmFilterZakaz
    Public ID_Sf As Integer

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmFilterZakaz_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dt As DataTable
        dt = Populate("select distinct ФИО_пациента from R_DMS where ID_SF=" + ID_Sf.ToString)
        cmbFIO.DataSource = dt
        cmbFIO.ValueMember = "ФИО_пациента"
        cmbFIO.DisplayMember = "ФИО_пациента"

    End Sub
End Class
