Imports System.Data.OLEDB
Imports System.Windows.Forms

Public Class frmFilterData
    Dim style As DataGridViewCellStyle = _
                    New DataGridViewCellStyle()
    Private adaptercombo As OleDbDataAdapter
    Private dscombo As New DataTable
    Public Source As String = ""
    Public Filter As String
    Public Per As Period

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If optPeriod.Checked Then
            Per.ID = cbPeriods.SelectedValue
        Else
            Per.dNach = dBegin.Value
            Per.dOkon = dEnd.Value
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
        
        modReport.dBegin = Per.dNach
        modReport.dEnd = Per.dOkon
        '(modReport.dBegin = "#1.1.1900#")        
    End Sub

    Public Sub SetFilter(ByRef frmL As frmList)
        'If optAll.Checked Then

        If optAll.Checked Then
            Exit Sub
        ElseIf optNoPay.Checked Then
            If frmL.TableName = "Q_SF_Beznal" Then
                frmL.Filter += " AND IsNull(Дата_последней_оплаты)=True"
                frmL.HeaderEnd = ". Неоплаченные"
            ElseIf frmL.TableName = "Q_SF_BeznalPr" Then
                frmL.Filter += " AND IsNull(Дата)=True"
                '    frmL.HeaderEnd = ". Неоплаченные"
            End If
        Else
            If frmL.TableName = "Q_SF_Beznal" Then
                frmL.Filter += " AND Дата_последней_оплаты>=" + FormatDateSQL(Per.dNach) + " AND Дата_последней_оплаты<=" + FormatDateSQL(Per.dOkon)
            ElseIf frmL.TableName = "Q_SF_BeznalPr" Then
                frmL.Filter += " AND Дата>=" + FormatDateSQL(Per.dNach) + " AND Дата<=" + FormatDateSQL(Per.dOkon)
                If optPeriod.Checked Then
                    frmL.HeaderEnd = ". За " + Per.Name
                Else
                    frmL.HeaderEnd = ". С " + FormatDateTime(Per.dNach, DateFormat.ShortDate) + " по " + FormatDateTime(Per.dOkon, DateFormat.ShortDate)
                End If
            End If
        End If
        'Me.Activate()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim dt As DataTable
        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from D_Periods", Connection)
        adaptercombo.Fill(dscombo)
        cbPeriods.DataSource = dscombo
        cbPeriods.ValueMember = "ID"
        cbPeriods.DisplayMember = "Наименование"
        'If Source = "СФ" Then
        '    Me.Text = "Выберите интервал последних дат оплат СФ"
        'Else
        '    'Select Case modReport.Templ
        '    '    Case "Реестр услуг.xlt"
        '    '        'dt = Populate("select * from D_MedUslug where ID IN (select ID_Uslugi from R_Uslugi)")
        '    '        dt = Populate("select MIN(Дата_оказания_услуги) As mindata, MAX(Дата_оказания_услуги) AS maxdata from R_Uslugi")
        '    '        Using reader As New DataTableReader(dt)
        '    '            If reader.Read Then
        '    '                If Not IsDBNull(reader("mindata")) Then dBegin.Value = reader("mindata")
        '    '                If Not IsDBNull(reader("maxdata")) Then dEnd.Value = reader("maxdata")
        '    '            End If
        '    '        End Using
        '    '        Me.Text = "Отчет по оказанным услугам"
        '    '        lblCaption.Text = "За период:"
        '    '        lblEnd.Visible = True
        '    '        lblBegin.Visible = True
        '    '        dEnd.Visible = True
        '    '    Case "Прейскурант.xlt"
        '    '        dEnd.Visible = False
        '    '        lblEnd.Visible = False
        '    '        lblBegin.Visible = False
        '    '        Me.Text = "Прейскурант"
        '    '        lblCaption.Text = "На дату:"

        '    'End Select
        'End If


    End Sub

    Private Sub optPeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPeriod.CheckedChanged
        cbPeriods.Enabled = optPeriod.Checked
    End Sub

    Private Sub optInterval_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optInterval.CheckedChanged
        dBegin.Enabled = optInterval.Checked
        dEnd.Enabled = optInterval.Checked
    End Sub
End Class
