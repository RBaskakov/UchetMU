Imports System.Windows.Forms

Public Class frmUsluga
    Public frmParent As frmList
    Private dsPodr As New DataTable
    Private dsUslugi As DataTable
    Private adPodr As New OleDb.OleDbDataAdapter
    Private adUslugi As New OleDb.OleDbDataAdapter
    Public ID_Year As Integer
    Public ID_Podr As Integer = 0
    Public ID_Uslugi As Integer = 0
    Private IsFormLoading As Boolean
    Private RsPodr, RsUslugi As OleDb.OleDbDataReader

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Debug.Print(ID_Uslugi)
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub PodrUnchoice()
        dsUslugi.Clear()
        cbUslugi.SelectedValue = -1
    End Sub

    Private Sub PodrChoice()
        Dim sql As String
        IsFormLoading = True
        If Not (dsUslugi Is Nothing) Then
            dsUslugi.Clear()
        Else
            dsUslugi = New DataTable
        End If
        cbPodr.SelectedValue = ID_Podr
        sql = "select * from Q_MedUslug_NoCalc Where ID_Podr=" + ID_Podr.ToString + " AND ID_Year=" + ID_Year.ToString + " ORDER BY Наименование"
        adUslugi = New OleDb.OleDbDataAdapter(sql, Connection)
        'dsPodr = New DataTable
        adUslugi.Fill(dsUslugi)
        cbUslugi.DataSource = dsUslugi
        cbUslugi.ValueMember = "ID"
        cbUslugi.DisplayMember = "Наименование"
        cbUslugi.SelectedValue = -1
        txtKodUslugi.Text = ""
        RsUslugi = readerBySQL(sql)
        IsFormLoading = False
    End Sub

    Private Sub frmUsluga_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        RsPodr.Close()
        RsUslugi.Close()
        'ID_Podr = -1
        'ID_Uslugi = -1
    End Sub

    Private Sub dlgUsluga_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim row As Integer
        Dim d As Date
        Dim sFilter As String

        With frmParent
            row = .dgwList.SelectedCells(0).RowIndex
            If .mTable(1) = "Q_Stom" Then
                ID_Podr = 26
            End If
            If IsColumnInDgw(.dgwList, "ID_Podr") Then
                If Not IsDBNull(.dgwList.Item("ID_Podr", row).Value) Then
                    ID_Podr = .dgwList.Item("ID_Podr", row).Value
                End If
            End If
            If Not IsDBNull(.dgwList.Item("ID_Uslugi", row).Value) Then
                ID_Uslugi = .dgwList.Item("ID_Uslugi", row).Value
            End If
            ID_Year = frmMDI.iYear
            If IsColumnInDgw(.dgwList, "Дата_оказания_услуги") Then
                If Not IsDBNull(.dgwList.Item("Дата_оказания_услуги", row).Value) Then
                    d = .dgwList.Item("Дата_оказания_услуги", row).Value
                    ID_Year = GetYearByDate(d)
                End If
            End If
            If IsColumnInDgw(.dgwList, "Дата_посещения") Then
                If Not IsDBNull(.dgwList.Item("Дата_посещения", row).Value) Then
                    d = .dgwList.Item("Дата_посещения", row).Value
                    ID_Year = GetYearByDate(d)
                End If
            End If

        End With
        RsPodr = readerBySQL("select * from D_Podr")
        RsUslugi = readerBySQL("select * from Q_MedUslug_NoCalc WHERE ID_Year=" + ID_Year.ToString)
        If RsUslugi.Read Then
            If ID_Podr > 0 Then
                sFilter = "ID_Podr=" + ID_Podr.ToString
            Else
                sFilter = ""
            End If
            If ID_Uslugi > 0 Then
                If sFilter <> "" Then
                    sFilter += " AND "
                End If
                sFilter += " ID=" + ID_Uslugi.ToString
            End If
            RsUslugi = readerBySQL("select * from Q_MedUslug_NoCalc WHERE " + sFilter)
            If RsUslugi.Read Then
                If ID_Uslugi > 0 Then
                    If sFilter = "" Then sFilter = "ID=" + ID_Uslugi.ToString
                    RsUslugi = readerBySQL("select * from Q_MedUslug_NoCalc WHERE " + sFilter)
                    If RsUslugi.Read Then
                        ID_Podr = RsUslugi("ID_Podr").Value
                    End If
                Else
                    sFilter = ""
                End If
            End If
        End If
        'commandBuilder = New OleDbCommandBuilder(adapter)
        'adaptercombo = Nothing
        'adPodr = New OleDbDataAdapter("select * from D_MedUslug Where ID IN (SELECT ID_Uslugi FROM D_MedUslug_Podr WHERE ID_Podr=" + i.ToString + ") ORDER BY Наименование", Connection)
        IsFormLoading = True
        adPodr = New OleDb.OleDbDataAdapter("select * from D_Podr ORDER BY Подразделение", Connection)
        'dsPodr = New DataTable
        adPodr.Fill(dsPodr)
        cbPodr.DataSource = dsPodr
        cbPodr.ValueMember = "ID"
        cbPodr.DisplayMember = "Подразделение"
        'cbPodr.SelectedValue = -1
        'cbUslugi.SelectedValue = -1
        IsFormLoading = False
        If ID_Podr > 0 Then
            cbPodr.SelectedValue = ID_Podr
        Else
            cbPodr.SelectedValue = -1
        End If
        If frmParent.mTable(1) = "Q_Stom" Then
            cbPodr.Enabled = False
            txtCodePodr.Enabled = False
        End If
        If ID_Uslugi > 0 Then
            If Not (cbUslugi.SelectedValue = ID_Uslugi) Then
                cbUslugi.SelectedValue = ID_Uslugi
            End If
            'cbUslugi_SelectedValueChanged(Nothing, Nothing)
        Else
            cbUslugi.SelectedValue = -1
            txtKodUslugi.Text = ""
        End If

    End Sub

    Private Sub txtCodePodr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodePodr.LostFocus
        If txtCodePodr.Text = "" Then Exit Sub
        RsPodr = readerBySQL("select * from D_Podr where Код_подразделения=" + txtCodePodr.Text)
        If RsPodr.Read Then
            If RsPodr.RecordsAffected = 1 Then
                ID_Podr = RsPodr("ID").Value
                PodrChoice()
                cbUslugi.SelectedValue = -1
            Else
                PodrUnchoice()
                txtCodePodr.Clear()
            End If
        Else
            PodrUnchoice()
            txtCodePodr.Clear()
        End If

    End Sub
    
    Private Sub cbPodr_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPodr.SelectedValueChanged
        If IsFormLoading Then Exit Sub
        ID_Podr = cbPodr.SelectedValue
        RsPodr = readerBySQL("select * from Q_Podr where ID=" + ID_Podr.ToString)
        If Not RsPodr.Read Then
            txtCodePodr.Text = ""
            Exit Sub
        End If
        txtCodePodr.Text = RsPodr("Код_подразделения").Value
        PodrChoice()
    End Sub

    Private Sub cbUslugi_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbUslugi.SelectedValueChanged
        If IsFormLoading Then Exit Sub
        If Not (cbUslugi.SelectedValue Is Nothing) Then
            ID_Uslugi = cbUslugi.SelectedValue
        Else
            ID_Uslugi = 0
            Exit Sub
        End If
        RsUslugi = readerBySQL("select * from Q_MedUslug_NoCalc WHERE ID" + ID_Uslugi.ToString)
        If Not RsUslugi.Read Then
            txtKodUslugi.Text = ""
            Exit Sub
        End If
        Me.txtKodUslugi.Text = RsUslugi("Код_услуги_по_подразделению").Value
    End Sub

    Private Sub txtKodUslugi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKodUslugi.LostFocus
        If txtKodUslugi.Text = "" Then Exit Sub
        If txtCodePodr.Text = "" Then Exit Sub
        IsFormLoading = True
        RsUslugi = readerBySQL("select * from Код_услуги_по_подразделению=" + txtKodUslugi.Text)
        If RsUslugi.Read Then
            ID_Uslugi = RsUslugi("ID").Value
            cbUslugi.SelectedValue = ID_Uslugi
        Else
            txtKodUslugi.Clear()
        End If
        IsFormLoading = False
    End Sub
End Class
