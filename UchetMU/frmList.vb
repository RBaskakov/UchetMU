Imports System.Data.OleDb
Imports System.Drawing
Imports System
Imports System.Windows.Forms

Public Class frmList
    Public ds As DataTable
    Public ds2 As New DataTable
    Private IsAddRow As Boolean = False
    Private dtLinkCols As New DataTable
    Private dtColAttr As New DataTable
    Private dtTablesAttr As New DataTable
    Public DSCombos As New Collection
    Public Combos As New Collection
    Public ComboAdapters As New Collection
    Public mTable As String()
    Private bindingSource As New BindingSource()
    Private bindingSource2 As New BindingSource()
    Public adapter As OleDb.OleDbDataAdapter
    Public adapter2 As OleDb.OleDbDataAdapter
    Private adapterAdd As OleDb.OleDbDataAdapter
    Public IsFitdgw As Boolean = True
    Private mNoEdit As Boolean()
    Public TableToUpdate As String()
    Public IsActivated As Boolean = False
    Public IsUpdateRow As Boolean = False
    Public IsAddRowByUser As Boolean = False
    Public Filter, Header As String
    Public LinkField As String
    Public FieldValues As New Collection
    Public SelectedRow As Integer()
    Private newid, newid2, LinkValue As Integer
    Public IsLinkingWithAnother As Boolean = False
    Public MenuInvoke As String
    Public RowDoubleClicked, RowDoubleClicked2 As Integer
    Public ID_VidSF As Integer
    Public IsCopy As Boolean = False
    Public HeaderEnd As String = ""
    Public WorkPer As Period


    Public Sub ApplyFilter()
        'Filter = Filter + " AND " + Flt
        Me.TableName = mTable(1)
        Me.Text = Me.Header + Me.HeaderEnd
        Me.frmList_Activated(Nothing, Nothing)
        'Me.Activate()
    End Sub

    Public Sub CreateLinkData(ByVal SenderName As String)
        Dim sTitle As String

        If Filter = "" Then Filter = "1=1"
        sTitle = ""
        Select Case SenderName
            Case "mnuSoldAmb"
                ID_VidSF = 1
                sTitle = "военнослужащие амбулаторно"
            Case "mnuSoldSt"
                ID_VidSF = 2
                sTitle = "военнослужащие по стационарам"
            Case "mnuDMS"
                ID_VidSF = 3
                sTitle = "ДМС"
            Case "mnuBeznalProekt"
                ID_VidSF = 5
                sTitle = "безнал-предварительные"
            Case "mnuBeznalFact"
                ID_VidSF = 4
                sTitle = "безнал-фактические"
            Case "mnuStom"
                ID_VidSF = 6
                sTitle = "стоматология"
            Case "rSoldAmb"
                Text = "ВС-амбулаторно: " + Text
                Exit Sub
            Case "rSoldSt"
                Text = "ВС-стационарно: " + Text
                Exit Sub
            Case "rDMS"
                Text = "ДМС: " + Text
                Exit Sub
            Case Else
                Exit Sub
        End Select
        FieldValues.Add(ID_VidSF, "ID_VidSF")
        Header = "Счет-фактуры: " + sTitle
        Text += "Счет-фактуры: " + sTitle
        Filter += " AND ID_VidSF=" + ID_VidSF.ToString

    End Sub

    Public Sub CreateLinkDataForForm(ByRef LinkForm As frmList)
        Dim i, row, NomLinkTable As Integer
        Dim d, dEnd As Date
        Dim dgw As DataGridView
        Dim tabName, s As String
        On Error GoTo Err_h
        'Dim dRow As DataRow

        'If LinkForm.dgwList2.SelectedCells.Count = 0 And LinkForm.dgwList2.Focus = True Then Exit Sub
        NomLinkTable = GetNomTableForLinkForm(LinkForm.TableName)
        If NomLinkTable = 1 Then
            tabName = LinkForm.TableName
            dgw = LinkForm.dgwList
        ElseIf NomLinkTable = 2 Then
            tabName = LinkForm.TableName2
            dgw = LinkForm.dgwList2
        End If
        'If dgw.SelectedCells.Count = 0 Then Exit Sub
        Filter = "1=1"
        Dim com As New OleDbCommand("select * from M_linkTablesData where Table='" + tabName + "'", Connection)
        Dim recL As OleDbDataReader = com.ExecuteReader()
        Do While recL.Read
            If recL("DataType").ToString.ToLower = "date" Then
                i = LinkForm.RowDoubleClicked2
                s = recL("LinkField")
                Text += " за " + dgw.Item(s, i).Value
                d = dgw.Item(s, i).Value
                dEnd = GetEndPeriod(d)
                s = recL("FieldForFilter")
                FieldValues.Add(d, s)
                Filter += " AND " + s + ">=" + FormatDateSQL(d) + " AND " + s + "<=" + FormatDateSQL(dEnd)
                'If LinkForm.TableName2 = "Q_CustomersByPeriod" Then
                'Else
                '    Filter += " AND Month(" + recL("FieldForFilter") + ")=" + d.Month.ToString + " AND Year(" + recL("FieldForFilter") + ")=" + d.Year.ToString
                'End If                
            End If
            If recL("DataType").ToString.ToLower = "int" Then
                If Text <> "" Then Text += ". "
                If LinkForm.TableName = "" Then
                    i = LinkForm.RowDoubleClicked2
                    Text += recL("Title") + ": " + LinkForm.dgwList.Item(recL("Title"), i)
                    Filter += " AND " + recL("LinkField").Value + "=" + LinkForm.dgwList2.Item(recL("LinkField"), i).ToString
                ElseIf LinkForm.TableName2 = "Q_MedUslug_NoCalc" Then
                    i = LinkForm.RowDoubleClicked2
                    row = LinkForm.dgwList.SelectedCells(0).RowIndex
                    Text = "Калькуляции на услугу: " + LinkForm.dgwList.Item("Код_подразделения", row).Value.ToString + "." + LinkForm.dgwList2.Item("Код_услуги_по_подразделению", i).Value.ToString + " (" + LinkForm.dgwList2.Item("Наименование", i).Value + ")"
                    LinkValue = LinkForm.dgwList2.Item("ID", i).Value
                    Filter += " AND ID_Uslugi=" + LinkForm.dgwList2.Item("ID", i).Value.ToString
                ElseIf LinkForm.TableName2 = "Q_CustomersByPeriod" Then
                    i = LinkForm.dgwList.SelectedCells(0).RowIndex
                    Filter += " AND ID_Customer=" + LinkForm.dgwList.Item("ID", i).Value.ToString
                    Text += recL("Title") + ": " + LinkForm.dgwList.Item(recL("Title"), i).Value
                    i = RowDoubleClicked2
                ElseIf LinkForm.TableName = "Q_SF" Or LinkForm.TableName Like "Q_SF_Beznal*" Or LinkForm.TableName = "Q_SF_DMS" Or LinkForm.TableName = "Q_SF_Stom" Then
                    i = LinkForm.RowDoubleClicked
                    LinkValue = LinkForm.dgwList.Item("ID", i).Value
                    Filter += " AND ID_SF=" + LinkForm.dgwList.Item("ID", i).Value.ToString
                    Text = "Спецификация счет-фактуры №" + LinkForm.dgwList.Item("Номер", i).Value + " от " + LinkForm.dgwList.Item("Дата", i).Value
                End If
                Debug.Print(Filter)
                FieldValues.Add(dgw.Item(recL("LinkField"), i), recL("FieldForFilter"))
                '      Filters.Add(recL("LinkField") + "=" + LinkForm.dgwList2.Item(recL("LinkField"), i).ToString, recL("FieldForFilter"))
                'dtDefault.Columns.Add(recL("FieldForFilter"), Type.GetType("System.Int32"))
                'i = LinkForm.dgwList.SelectedCells(0).RowIndex
            End If
        Loop
        recL.Close()
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.CreateLinkData")
    End Sub

    Public Function IsNoEdit(ByVal NomTable As Integer) As Boolean
        Return mNoEdit(NomTable)
    End Function

    Public Sub SetNoEdit(ByVal NomTable As Integer, ByVal Value As Boolean)
        Dim dgw As DataGridView
        Dim i As Integer

        If Value = False Then Exit Sub
        'If mNoEdit.Length = 0 Then ReDim mNoEdit(3)
        mNoEdit(NomTable) = Value
        dgw = GetdgwByNom(NomTable)

        For i = 1 To dgw.Columns.Count - 1
            If dgw.Columns(i).Name <> "Dummy" Then
                dgw.Columns(i).ReadOnly = Value
            End If
        Next i

    End Sub

    Property TableName() As String
        Get
            Return mTable(1)
        End Get
        Set(ByVal Value As String)
            Dim commandBuilder As OleDbCommandBuilder
            Dim i As Integer
            Dim sql, s As String
            On Error GoTo Err_hand

            IsFitdgw = False
            mTable(1) = Value
            ds = New DataTable
            bindingSource = New BindingSource
            ds.Locale = System.Globalization.CultureInfo.InvariantCulture
            If Filter = "" Then Filter = "1=1"
            If Connection.State = ConnectionState.Closed Then Connection.Open()

            If mTable(1) Like "Q_SF_Beznal*" Then
                Dim frm As New frmFilterData
                'frm.lblCaption.Text = "Интервал дат"
                If mTable(1) = "Q_SF_Beznal" Then
                    frm.Source = "СФ"
                    frm.Text = "Выберите интервал дат последней оплаты СФ"
                ElseIf mTable(1) = "Q_SF_BeznalPr" Then
                    frm.Source = "СФ проект"
                    frm.Text = "Выберите интервал дат СФ"
                    frm.optNoPay.Text = "Без дат"
                End If
                frm.OK_Button.Text = "OK"
                If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    frm.SetFilter(Me)
                    Me.Text = Me.Header + Me.HeaderEnd
                Else
                    mTable(1) = ""
                    Exit Property
                End If
            End If

            GetLinkColumns(1)
            GetColumnsAttr(1)
            sql = "select * from " + mTable(1) + " where " + Filter
            s = GetSortField(1)
            If s <> "" Then
                sql += " ORDER By " + s
            End If
            Debug.Print(sql)
            adapter = New OleDb.OleDbDataAdapter(sql, Connection)
            commandBuilder = New OleDbCommandBuilder(adapter)

            ds.Clear()
            adapter.Fill(ds)

            dgwList.AutoGenerateColumns = Not (dgwList.Columns.Count > 0)
            bindingSource.DataSource = ds
            dgwList.DataSource = bindingSource
            For i = 0 To dgwList.Columns.Count - 1
                If (dgwList.Columns(i).Name Like "ID*") Then dgwList.Columns(i).Visible = False
            Next i
            dgwList.AutoGenerateColumns = False
            If (mTable(1) Like "Q*") Then
                TableToUpdate(1) = GetTableToUpdate(mTable(1))
                'TableToUpdate(1) = mTable(1).Replace("Q_", "R_")
                GetEditColumns(1)
                CreateLinkColumns(1)
            ElseIf (mTable(1) Like "R_*") Then
                TableToUpdate(1) = mTable(1)
                GetEditColumns(1)
                CreateLinkColumns(1)
            Else
                TableToUpdate(1) = mTable(1)
                GetEditColumns(1)
                CreateLinkColumns(1)
            End If
            If mNoEdit(1) = False Then
                If Connection.State = ConnectionState.Closed Then Connection.Open()
            End If
            If Not dgwList.Columns.Contains("Dummy") Then dgwList.Columns.Add("Dummy", "")
            dgwList.Columns(dgwList.Columns.Count - 1).ReadOnly = True
            SetNoEdit(1, mNoEdit(1))
            If dgwList.Columns.Count > 5 Then
                If dgwList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                    dgwList.Columns(dgwList.Columns.Count - 1).Width = 1
                ElseIf dgwList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill Then
                    dgwList.Columns(dgwList.Columns.Count - 1).FillWeight = 0.05
                End If
            End If
            For i = 0 To dgwList.Columns.Count - 1
                dgwList.Columns(i).HeaderText = dgwList.Columns(i).HeaderText.Replace("_", " ")
                If i > dgwSumma.Columns.Count - 1 Then
                    dgwSumma.Columns.Add(i.ToString, "")
                End If
                dgwSumma.Columns(i).Visible = dgwList.Columns(i).Visible
            Next i
            If dgwSumma.RowCount = 0 Then dgwSumma.Rows.Add()
            dgwSumma.Visible = IsAnySumColumn(1)
            If IsAnySumColumn(1) Then
                ReCalcAllSum(1)
                IsFitdgw = True
                Exit Property
            End If
            IsFitdgw = True
            Exit Property
Err_hand:
            MsgBox("Ошибка в процедуре frmList.TableName :" & Err.Description)
        End Set
    End Property

    Property TableName2() As String
        Get
            Return mTable(2)
        End Get
        Set(ByVal Value As String)
            On Error GoTo Err_hand
            Dim commandBuilder As OleDbCommandBuilder
            Dim i As Integer
            Dim s, sql As String

            mTable(2) = Value
            If mTable(2) = "" Then
                IsFitdgw = True
                'SplitContainer1.Visible = False
                SplitContainer.Dock = DockStyle.Fill
                Exit Property
            Else
                IsFitdgw = False
            End If
            ds2.Locale = System.Globalization.CultureInfo.InvariantCulture
            sql = "select * from " + mTable(2)
            s = GetSortField(2)
            If s <> "" Then
                sql += " ORDER By " + s
            End If
            adapter2 = New OleDbDataAdapter(sql, Connection)
            adapter2.ContinueUpdateOnError = True
            commandBuilder = New OleDbCommandBuilder(adapter2)

            If Connection.State = ConnectionState.Closed Then Connection.Open()
            ds2.Clear()
            adapter2.Fill(ds2)

            dgwList2.AutoGenerateColumns = True
            bindingSource2.DataSource = ds2
            bindingSource2.Filter = "1=0"
            dgwList2.DataSource = bindingSource2
            For i = 0 To dgwList2.Columns.Count - 1
                If (dgwList2.Columns(i).Name Like "ID*") Then dgwList2.Columns(i).Visible = False
            Next i
            dgwList2.AutoGenerateColumns = False
            If (mTable(2) Like "Q*") Then
                TableToUpdate(2) = GetTableToUpdate(mTable(2))
                CreateLinkColumns(2)
                GetEditColumns(2)
            ElseIf (mTable(2) Like "R_*") Then
                TableToUpdate(2) = mTable(2)
                CreateLinkColumns(2)
                GetEditColumns(2)
            Else
                TableToUpdate(2) = mTable(2)
                CreateLinkColumns(2)
            End If
            If mNoEdit(2) = False Then
                dgwList2.Columns.Add("Dummy", "")
            End If
            dgwList2.Columns(dgwList2.Columns.Count - 1).ReadOnly = True
            If dgwList2.Columns.Count > 5 Then
                If dgwList2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                    dgwList2.Columns(dgwList2.Columns.Count - 1).Width = 1
                ElseIf dgwList2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill Then
                    dgwList2.Columns(dgwList2.Columns.Count - 1).FillWeight = 0.05
                End If
            End If

            For i = 0 To dgwList2.Columns.Count - 1
                dgwList2.Columns(i).HeaderText = dgwList2.Columns(i).HeaderText.Replace("_", " ")
                If i > dgwSumma2.Columns.Count - 1 Then
                    dgwSumma2.Columns.Add(i.ToString, "")
                End If
                dgwSumma2.Columns(i).Visible = dgwList2.Columns(i).Visible
            Next i
            SetNoEdit(2, mNoEdit(2))
            If IsAnySumColumn(2) Then
                dgwSumma2.Rows.Add()
                ReCalcAllSum(2)
            End If
            'Split2.Panel2Collapsed = Not IsAnySumColumn(2)

            If IsAnyReCalcColumn(2) Then
                IsFitdgw = True
                Exit Property
            End If
            IsFitdgw = True
            Exit Property
Err_hand:
            MsgBox("Ошибка в процедуре frmList.TableName2 :" & Err.Description)
        End Set
    End Property

    'Public Sub RefreshData()

    'End Sub

    Private Function IsColumnReadyToFill(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()
        Dim row As Integer

        dgw = GetdgwByNom(NomTable)
        s = dgw.Columns(ColIndex).Name
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "'")
        If dRows.Length = 0 Then Return True
        If IsDBNull(dRows(0)("FillAfter")) Then Return True
        s = CStr(dRows(0)("FillAfter"))
        If dgw.SelectedCells.Count = 0 Then Return True
        row = dgw.SelectedCells(0).RowIndex
        If IsDBNull(dgw.Item(s, row).Value) Then
            MsgBox("Сначала необходимо заполнить поле '" + dgw.Columns(s).HeaderText.ToString + "'")
            Return False
        End If
        If CStr(dgw.Item(s, row).Value) = "" Or CStr(dgw.Item(s, row).Value) = "0" Then
            MsgBox("Сначала необходимо заполнить поле '" + dgw.Columns(s).HeaderText.ToString + "'")
            Return False
        End If
        Return True
    End Function

    Private Function IsColumnRecalc(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()

        dgw = GetdgwByNom(NomTable)
        s = dgw.Columns(ColIndex).Name
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "' And IsRecalcColumn=True")
        Return (dRows.Length > 0)

    End Function

    '    Private Function GetKeyField(ByVal NomTable As Integer, ByVal ColIndex As Integer) As String
    '        Dim s As String
    '        Dim dgw As DataGridView
    '        Dim dRows As DataRow()
    '        On Error GoTo Err_h

    '        dgw = GetdgwByNom(NomTable)
    '        s = dgw.Columns(ColIndex).Name
    '        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "'")
    '        If dRows.Length = 0 Then Return ""
    '        If IsDBNull(dRows(0)("KeyField")) Then Return ""
    '        Return dRows(0)("KeyField").Value
    'Err_h:
    '        Return ""
    '    End Function

    Private Function GetSortField(ByVal NomTable As Integer) As String
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND IsSortColumn=True")
        If dRows.Length = 0 Then
            Return ""
        Else
            Return dRows(0)("ColumnName")
        End If

Err_h:
        Return ""
    End Function

    Public Function IsTable(ByVal NomTable As Integer, ByVal Prop As String) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        dRows = dtTablesAttr.Select("Table='" + mTable(NomTable) + "' AND " + Prop + "=True")
        Return (dRows.Length > 0)
Err_h:
        Return False
    End Function

    Private Function IsLinkTextColumn(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        s = dgw.Columns(ColIndex).Name
        dRows = dtLinkCols.Select("Table='" + mTable(NomTable) + "' AND Column='" + s + "' AND TypeLink=0")
        Return (dRows.Length > 0)
Err_h:
        Return False
    End Function

    Private Function IsColumn(ByVal ColIndex As Integer, ByVal NomTable As Integer, ByVal Prop As String) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        s = dgw.Columns(ColIndex).Name
        'If Not (Prop Like "*=*") Then Prop += "=True"
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "' AND " + Prop + "=True")
        Return (dRows.Length > 0)
Err_h:
        Return False
    End Function

    'Private Function IsColumnObligatory(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
    '    Dim s As String
    '    Dim dgw As DataGridView
    '    Dim dRows As DataRow()

    '    dgw = GetdgwByNom(NomTable)
    '    s = dgw.Columns(ColIndex).Name
    '    dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "' And IsObligatory=True")
    '    Return (dRows.Length > 0)

    'End Function

    'Private Function IsColumnDataUpdateWhenEnter(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
    '    Dim s As String
    '    Dim dgw As DataGridView
    '    Dim dRows As DataRow()

    '    dgw = GetdgwByNom(NomTable)
    '    s = dgw.Columns(ColIndex).Name
    '    dRows = dtColAttr.Select("Table='" + TableToUpdate(NomTable) + "' AND ColumnName='" + s + "' And IsUpdateWhenEnter=True")
    '    Return (dRows.Length > 0)

    'End Function

    Private Function IsAnySumColumn(ByVal NomTable As Integer) As Boolean
        Dim dgw As DataGridView
        Dim dRows As DataRow()

        dgw = GetdgwByNom(NomTable)
        's = dgw.Columns(ColIndex).Name
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND IsSumColumn=True")
        Return (dRows.Length > 0)

    End Function

    Private Function IsAnyReCalcColumn(ByVal NomTable As Integer) As Boolean
        Dim dgw As DataGridView
        Dim dRows As DataRow()

        dgw = GetdgwByNom(NomTable)
        's = dgw.Columns(ColIndex).Name
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND IsRecalcColumn=True")
        Return (dRows.Length > 0)

    End Function

    Private Function IsSumColumn(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
        Dim s As String
        Dim dgw As DataGridView
        Dim dRows As DataRow()

        dgw = GetdgwByNom(NomTable)
        s = dgw.Columns(ColIndex).Name
        dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "' And IsSumColumn=True")
        Return (dRows.Length > 0)

    End Function

    'Public Function Populate(ByVal sqlString As String) As DataTable
    '    Dim dt As DataTable
    '    dt = modBas.Populate(sqlString)
    '    Return dt
    'End Function

    Private Sub GetEditColumns(ByVal NomTable As Integer)
        Dim dt As DataTable
        Dim i As Integer
        Dim dgw As DataGridView
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        For i = 0 To dgw.Columns.Count - 1
            dgw.Columns(i).ReadOnly = False
        Next i
        'If EditColumns.Count > 0 Then EditColumns.Clear()
        dt = Populate("select * from M_Columns where Table='" + mTable(NomTable) + "' AND IsNoEdit=True")
        Using reader As New DataTableReader(dt)
            Do While reader.Read
                'EditColumns.Add(reader("Column"), reader("Column"))
                If IsColumnInDgw(dgw, reader("ColumnName")) Then
                    Debug.Print(reader("ColumnName").ToString)
                    dgw.Columns(reader("ColumnName").ToString).ReadOnly = True
                End If
            Loop
        End Using
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.GetEditColumns")
    End Sub

    Private Sub CreateLinkColumns(ByVal NomTable As Integer)
        Dim dt As DataTable
        Dim col, tab, s, s2 As String
        Dim i, pos As Integer
        Dim dgw As DataGridView
        On Error GoTo Err_h

        tab = mTable(NomTable)
        s = ""
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        dgw = GetdgwByNom(NomTable)
        dt = Populate("select * from M_LinkData where Table='" + tab + "' ORDER BY ID")
        Using reader As New DataTableReader(dt)
            Do While reader.Read
                If Not IsDBNull(reader("ID_Column")) Then
                    s = reader("ID_Column").ToString
                    s2 = reader("LinkTable").ToString
                    If IsColumnInDgw(dgw, s) And Not IsColumnInDgw(dgw, s2) And reader("TypeLink") = 1 Then
                        i = dgw.Columns(reader("ID_Column")).Index
                        col = dgw.Columns(reader("ID_Column")).Name
                        dgw.Columns.Insert(i + 1, CreateComboBox(IIf(IsDBNull(reader("LinkTable")), "", reader("LinkTable")), reader("ID_Column"), reader("ColumnInLinkTable")))
                        dgw.Columns(i).HeaderText = col.ToString.Replace("_", " ")
                        dgw.Columns(i + 1).HeaderText = reader("Column").ToString.Replace("_", " ")
                    ElseIf reader("TypeLink") = 0 Then
                        CreateDSCombo(IIf(IsDBNull(reader("LinkTable")), "", reader("LinkTable")), reader("ID_Column"), reader("ColumnInLinkTable"))
                        If Not dgw.Columns(reader("Column")) Is Nothing Then dgw.Columns(reader("Column")).ReadOnly = True
                    End If
                End If
            Loop
        End Using
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.CreateLinkColumns")
    End Sub

    Public Sub RefreshLinkColumnData(ByVal KeyCol As Integer, ByVal ColVal As Integer, ByVal NomTable As Integer)
        Dim i, j As Integer
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        For i = 0 To dgw.Rows.Count - 1
            If Not IsDBNull(dgw.Item(KeyCol, i).Value) Then
                j = dgw.Item(KeyCol, i).Value
                IsUpdateRow = True
                dgw.Item(ColVal, i).Value = j
                IsUpdateRow = False
            End If
        Next i
    End Sub

    Public Sub FillRecalcColumns(ByVal NomTable As Integer, ByVal Row As Integer)
        Dim dgw As DataGridView
        Dim q, col, str As String
        Dim i, ID As Integer
        Dim s As Single
        Dim dRow, dRow2 As DataRow()
        Dim dt As DataTable
        Dim dt2 As New DataTable
        'Dim adapter As OleDbDataAdapter
        On Error GoTo Err_Hand

        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        If Not IsActivated Then Exit Sub
        dgw = GetdgwByNom(NomTable)
        dt = GetDataTableByNom(NomTable)
        'If Not IsObjectExist(q) Then Exit Sub
        If Row > dgw.Rows.Count - 1 Then Exit Sub
        If IsDBNull(dgw.Item("ID", Row).Value) Or dgw.Item("ID", Row).Value Is Nothing Then Exit Sub
        ID = dgw.Item("ID", Row).Value
        dRow = dt.Select("ID=" + ID.ToString)
        'dRow(0).Delete()
        Dim command As New OleDbCommand("SELECT * FROM " & mTable(NomTable) + " WHERE ID=" + ID.ToString, Connection)
        Dim adap As New OleDbDataAdapter()
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        adap.SelectCommand = command
        adap.Fill(dt2)
        dRow2 = dt2.Select("ID=" + ID.ToString)
        CopyDataRow(dRow2(0), dRow(0))
        RefreshAllComboBox(NomTable, Row)
        For i = 1 To dgw.Columns.Count - 1
            ReCalcSum(i, NomTable)
        Next i
        IsUpdateRow = False
        dgw.Refresh()

        Exit Sub
Err_Hand:
        ErrMess("Ошибка в процедуре frmList.FillLRecalcColumns: " & Err.Description)
    End Sub

    Public Function CreateComboBox(ByVal sTable As String, ByVal ID_Column As String, ByVal NameColumn As String) As DataGridViewComboBoxColumn
        Dim combo As New DataGridViewComboBoxColumn()
        Dim adaptercombo As OleDbDataAdapter
        Dim dscombo As New DataTable

        If sTable = "" Then
            combo.ValueMember = "ID"
            combo.DisplayMember = NameColumn
            combo.Name = NameColumn
            Combos.Add(combo, NameColumn)
            Return combo
        End If
        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from " + sTable + " order by " + NameColumn, Connection)
        adaptercombo.Fill(dscombo)
        DSCombos.Add(dscombo, sTable)
        ComboAdapters.Add(adaptercombo, sTable)
        combo.DataSource = DSCombos(sTable)
        combo.ValueMember = "ID"
        combo.DisplayMember = NameColumn
        combo.Name = sTable
        Combos.Add(combo, sTable)

        Return combo
    End Function

    Public Sub CreateDSCombo(ByVal sTable As String, ByVal ID_Column As String, ByVal NameColumn As String)
        Dim combo As New DataGridViewComboBoxColumn()
        Dim adaptercombo As OleDbDataAdapter
        Dim dscombo As New DataTable

        dscombo.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adaptercombo = Nothing
        adaptercombo = New OleDbDataAdapter("select * from " + sTable + " order by " + NameColumn, Connection)
        adaptercombo.Fill(dscombo)
        DSCombos.Add(dscombo, sTable)
        'ComboAdapters.Add(adaptercombo, sTable)

    End Sub

    Public Function CreateTextBox(ByVal NameColumn As String) As DataGridViewTextBoxColumn
        Dim txt As New DataGridViewTextBoxColumn()

        txt.Name = NameColumn
        txt.ValueType = System.Type.GetType("String")
        txt.HeaderText = NameColumn

        Return txt
    End Function

    Private Sub frmList_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim i, row As Integer

        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        frmMDI.mnuFilter.Visible = IsFilterVisible()
        frmMDI.mnuAdmin.Visible = (frmMDI.MdiChildren.Length = 0)
        frmMDI.mnuSFReport.Visible = (Mid(TableName, 1, 4) = "Q_SF")
        frmMDI.mnuAkt.Visible = (TableName = "Q_SF_DMS" Or TableName = "Q_SF_Stom")
        frmMDI.mnuZakaz.Visible = (TableName = "Q_SF_DMS" Or TableName = "Q_SF_Stom")

        If dgwList.SelectedCells.Count > 0 Then
            FillRecalcColumns(1, dgwList.SelectedCells(0).RowIndex)
        End If
        dgwSumma.Visible = IsAnySumColumn(1)
        For i = 0 To dgwList.Columns.Count - 1
            If dgwList.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                For row = 0 To dgwList.Rows.Count - 1
                    RefreshComboBox(1, i, row)
                Next row
                dgwList.Columns(i - 1).Visible = False
                dgwSumma.Columns(i - 1).Visible = False
            End If
        Next i
        For i = 0 To dgwList2.Columns.Count - 1
            If dgwList2.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                For row = 0 To dgwList2.Rows.Count - 1
                    RefreshComboBox(2, i, row)
                Next row
                dgwList2.Columns(i - 1).Visible = False
            End If
        Next i
        dgwList2.Visible = mTable(2) <> ""
        If mTable(2) <> "" Then
            If dgwList2.SelectedCells.Count > 0 Then
                'FillRecalcColumns(2, dgwList2.SelectedCells(0).RowIndex)
            End If
            dgwSumma2.Visible = IsAnySumColumn(2)
            If dgwSumma2.Visible Then
                ReCalcAllSum(2)
            End If
        End If
        IsActivated = True

    End Sub

    Private Sub dgwList_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgwList.CellBeginEdit
        On Error GoTo Err_Hand
        Dim dgw As DataGridView
        Dim Cell As DataGridViewCell
        Dim col As String
        Dim ScrollPos As Integer

        If Not IsActivated Then Exit Sub
        If Not IsColumnReadyToFill(CInt(e.ColumnIndex), 1) Then
            e.Cancel = True
            Exit Sub
        End If
        dgw = GetdgwByNom(1)
        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" And Not IsCopy Then
            If MsgBox("Редактировать данное поле?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                e.Cancel = True
            End If
            Exit Sub
        End If
        If dgw.Columns(CInt(e.ColumnIndex)).ValueType Is Nothing Then
            cldr1.Visible = False
            Exit Sub
        End If
        If dgw.Columns(CInt(e.ColumnIndex)).ValueType.Name <> "DateTime" Then
            cldr1.Visible = False
            Exit Sub
        End If
        If cldr1.Visible Then Exit Sub
        If dgw.AllowUserToAddRows And e.RowIndex = dgw.Rows.Count - 1 Then Exit Sub
        With cldr1
            .Left = CalendarLeft(1, e.ColumnIndex)
            .Top = CalendarTop(1, e.RowIndex)
            If e.RowIndex < dgw.Rows.Count - 1 And Not (dgw.Item(e.ColumnIndex, e.RowIndex).Value Is System.DBNull.Value) Then
                .SetDate(dgw.Item(e.ColumnIndex, e.RowIndex).Value)
            Else
                .SetDate(Today)
            End If
            .Visible = True
        End With
        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "dgwList_CellBeginEdit")

    End Sub

    Public Sub SaveData()
        Dim i As Integer
        On Error GoTo Err_h

        If Not (mTable(2) Like "Q*") And mTable(2) <> "" Then
            adapter2.Update(ds2)
            'Updatedgw(2)
        ElseIf IsTable(1, "IsBatchUpdate") Then
            adapter.Update(ds)
        End If

        Exit Sub
Err_h:
        MsgBox("Ошибка при сохранении данных: " + Err.Description)
    End Sub

    Public Sub SaveColumnWidthToSettings(ByVal NomTable As Integer)
        Dim i As Integer
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        For i = 0 To dgw.Columns.Count - 1
            If dgw.Columns(i).Visible Then
                If dgw.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                    SaveSetting("MedService", MenuInvoke & "_T1", dgw.Columns(i).Name & "_Width", dgw.Columns(i).Width)
                ElseIf dgw.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill Then
                    SaveSetting("MedService", MenuInvoke & "_T1", dgw.Columns(i).Name & "_Width", dgw.Columns(i).FillWeight)
                End If
            End If
        Next i
    End Sub

    Public Sub GetColumnWidthFromSettings(ByVal NomTable As Integer)
        On Error GoTo Err_h
        Dim i As Integer
        Dim width As Single
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        For i = 0 To dgw.Columns.Count - 1
            If dgw.Columns(i).Visible Then
                width = GetSetting("MedService", MenuInvoke & "_T1", dgw.Columns(i).Name & "_Width", 0)
                If width > 0 Then
                    If dgw.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                        dgw.Columns(i).Width = width
                    ElseIf dgw.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill Then
                        dgw.Columns(i).FillWeight = width
                    End If
                End If
            End If
        Next i
        Exit Sub
Err_h:
        Err.Clear()
    End Sub

    Private Sub frmList_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        On Error GoTo Err_Hand
        Filter = ""
        frmMDI.mnuFilter.Visible = IsFilterVisible()
        frmMDI.mnuAdmin.Visible = (frmMDI.MdiChildren.Length = 0)
        frmMDI.mnuSFReport.Visible = (frmMDI.MdiChildren.Length > 0)

        'frmMDI.mnuFileDB.Visible = (frmMDI.MdiChildren.Length = 0)
        'frmMDI.mnuYear.Visible = (frmMDI.MdiChildren.Length = 0)
        SaveData()
        'frmMDI.mnuLinkWithAnother.Visible = (frmMDI.MdiChildren.Length > 0)
        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "frmList_Disposed")
    End Sub

    Private Sub dgwList_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList.CellDoubleClick
        RowDoubleClicked = e.RowIndex
        If RowDoubleClicked = -1 Then Exit Sub
        If IsLinkTextColumn(e.ColumnIndex, 1) Then
            If Not IsColumnReadyToFill(CInt(e.ColumnIndex), 1) Then
                Exit Sub
            End If
            Dim frm As New frmCombo
            Dim dRows As DataRow()
            Dim ColName As String
            ColName = dgwList.Columns(e.ColumnIndex).Name
            dRows = dtLinkCols.Select("Column='" + ColName + "' AND Table='" + mTable(1) + "'")
            If dRows.Length = 0 Then Exit Sub
            frm.ColumnName = dRows(0)("ColumnInLinkTable")
            frm.STable = dRows(0)("LinkTable")
            frm.DataSource = DSCombos(frm.STable)
            frm.ColumnTitle = dgwList.Columns(e.ColumnIndex).HeaderText
            'frm.frmParent = Me
            frm.SelectedValue = dgwList.Item(e.ColumnIndex - 1, e.RowIndex).Value
            frm.ShowDialog()
            If frm.DialogResult = Windows.Forms.DialogResult.OK Then
                UpdateValueInDgw(Me, 1, e.RowIndex, dgwList.Columns(e.ColumnIndex - 1).Name, frm.combo.SelectedValue)
                'dgwList.Item("ID_Uslugi", e.RowIndex).Value = frm.ID_Uslugi
                FillRecalcColumns(1, e.RowIndex)
                dgwList.Item(e.ColumnIndex, e.RowIndex).Selected = True
            End If
        ElseIf dgwList.Columns(e.ColumnIndex).Name = "Услуга" Or dgwList.Columns(e.ColumnIndex).Name = "Код_услуги_по_номенклатуре" Then
            If Not IsColumnReadyToFill(CInt(e.ColumnIndex), 1) Then
                Exit Sub
            End If
            Dim frm As New frmUsluga
            frm.frmParent = Me
            frm.ShowDialog()
            If frm.DialogResult = Windows.Forms.DialogResult.OK Then
                UpdateValueInDgw(Me, 1, e.RowIndex, "ID_Uslugi", frm.ID_Uslugi)
                'dgwList.Item("ID_Uslugi", e.RowIndex).Value = frm.ID_Uslugi
                If IsColumnInDgw(dgwList, "ID_Podr") Then
                    UpdateValueInDgw(Me, 1, e.RowIndex, "ID_Podr", frm.ID_Podr)
                    'dgwList.Item("ID_Podr", e.RowIndex).Value = frm.ID_Podr
                    '    RefreshComboBox(1, dgwList.Columns("D_Podr").Index, e.RowIndex)
                End If
                FillRecalcColumns(1, e.RowIndex)
                dgwList.Item(e.ColumnIndex, e.RowIndex).Selected = True
            End If
        ElseIf TableToUpdate(1) Like "R_SF*" Then
            'Dim ID_VidSF As Integer
            'ID_VidSF = FieldValues("ID_VidSF")
            Select Case FieldValues("ID_VidSF")
                Case 1
                    frmMDI.ViewForm("rSoldAmb", Me)
                Case 2
                    frmMDI.ViewForm("rSoldSt", Me)
                Case 3
                    frmMDI.ViewForm("rDMS", Me)
                Case 4
                    frmMDI.ViewForm("rSpecSF", Me)
                Case 5
                    frmMDI.ViewForm("rSpecSF", Me)
                Case 6
                    frmMDI.ViewForm("rStom", Me)
            End Select

        End If
    End Sub

    Private Sub dgwList_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList.CellEndEdit
        On Error GoTo Err_Hand
        Dim dgw As DataGridView

        dgw = GetdgwByNom(1)
        If dgw.Visible = False Then Exit Sub
        If e.ColumnIndex = 0 Then Exit Sub
        If IsDBNull(dgwList.Item("ID", e.RowIndex).Value) And newid > 0 Then
            dgwList.Item("ID", e.RowIndex).Value = newid
            newid = 0
        End If
        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then
            dgw.Item(e.ColumnIndex - 1, e.RowIndex).Value = dgw.Item(e.ColumnIndex, e.RowIndex).Value
            'Updatedgw(1)
        ElseIf dgwList.Columns(e.ColumnIndex).ValueType Is Nothing Then
            cldr1.Visible = False
            Exit Sub
        ElseIf dgwList.Columns(e.ColumnIndex).ValueType.ToString = "DateTime" Then
            cldr1.Visible = False
        End If
        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "dgwList_CellEndEdit")
    End Sub


    Private Function CalendarLeft(ByVal NomTable As Integer, ByVal ColumnIndex As Integer) As Integer
        Dim i As Integer
        Dim dgw As DataGridView
        Dim iLeft As Integer = 0

        dgw = GetdgwByNom(NomTable)
        iLeft = 0
        iLeft += dgw.RowHeadersWidth
        For i = 1 To ColumnIndex
            If dgw.Columns(i).Visible Then iLeft += dgw.Columns(i).Width
        Next i
        If NomTable = 1 Then
            If iLeft + cldr1.Width > dgw.Width Then iLeft = iLeft - dgw.Columns(ColumnIndex).Width - cldr1.Width
        ElseIf NomTable = 2 Then
            If iLeft + cldr1.Width > dgw.Width Then iLeft = iLeft - dgw.Columns(ColumnIndex).Width - cldr1.Width
        End If

        Return iLeft
    End Function

    Private Function CalendarTop(ByVal NomTable As Integer, ByVal RowIndex As Integer) As Integer
        Dim iTop As Integer = 0
        Dim dgw As DataGridView
        'dgw = GetdgwByNom(NomTable)
        'iTop += dgw.RowHeadersWidth
        'For i = 1 To RowIndex
        '    iTop += dgw.Rows(i).Height
        'Next i
        'If iTop > dgw.Height Then iTop = dgw.Height / 2
        'dgwList.Height
        dgw = GetdgwByNom(NomTable)
        If NomTable = 1 Then
            iTop = Windows.Forms.Cursor.Position.Y - 90
            If iTop + cldr1.Height > dgw.Height Then iTop = dgw.Height - cldr1.Height
        ElseIf NomTable = 2 Then
            iTop = Windows.Forms.Cursor.Position.Y - 80 - SplitContainer.Panel1.Height - IIf(dgwSumma.Visible, dgwSumma.Height, 0)
            If iTop + cldr2.Height > dgw.Height Then iTop = dgw.Height - cldr2.Height
        End If

        Return iTop
    End Function

    Private Sub dgwList_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList.CellValueChanged
        Dim i As Integer
        Dim col As String
        Dim reader As OleDbDataReader
        On Error GoTo Err_Hand
        If Not IsActivated Then Exit Sub
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        Debug.Print(dgwList.Columns(e.ColumnIndex).Name)
        If IsUpdateRow Then
            ReCalcSum(e.ColumnIndex, 1)
            Exit Sub
        End If
        If e.RowIndex <= -1 Then
            Exit Sub
        End If
        'If mTable(2) = "" And Not IsTable(1, "mnuAdd") Then Exit Sub
        If IsTable(1, "IsBatchUpdate") Then Exit Sub
        If dgwList.Columns(e.ColumnIndex).ReadOnly Then Exit Sub
        UpdateCellADO(1, e.ColumnIndex, e.RowIndex)
        'UpdateDependCell(mTable(1), e.ColumnIndex, e.RowIndex, dgwList)
        FillRecalcColumns(1, e.RowIndex)

        ReCalcSum(e.ColumnIndex, 1)

        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "dgwList_CellValueChanged")
    End Sub

    Private Sub dgwList_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgwList.ColumnWidthChanged
        Dim i, width As Integer
        If Not IsActivated Then Exit Sub
        'If dgwSumma.Visible Then
        width = 50
        For i = 0 To dgwList.Columns.Count - 1
            If dgwList.Columns(i).Visible Then
                dgwSumma.Columns(i).Width = dgwList.Columns(i).Width
                width += dgwList.Columns(i).Width
                'dgwSumma.Columns(i).DefaultCellStyle = style
                'width += dgwList.Columns(i).Width
            End If
        Next i
        'Split1.Width = width
        'End If
    End Sub

    Private Sub dgwList_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgwList.DataError
        Err.Clear()
    End Sub

    Public Sub Fitdgw()
        Dim i, width As Integer
        Dim style As DataGridViewCellStyle = _
                        New DataGridViewCellStyle()
        Dim f As Font
        On Error GoTo Err_Hand

        If Me.Visible = False Or IsFitdgw = False Or IsActivated = False Then Exit Sub
        style = dgwSumma.DefaultCellStyle
        style.BackColor = Me.BackColor
        f = New Font(style.Font, Drawing.FontStyle.Bold)
        style.Font = f
        IsFitdgw = False
        dgwSumma.RowHeadersWidth = dgwList.RowHeadersWidth
        If dgwList.Columns.Count <> dgwSumma.Columns.Count Then Exit Sub
        width = 50
        If dgwList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill Then
            dgwList.Width = Me.Width - 25
        End If
        dgwSumma.Width = dgwList.Width
        For i = 0 To dgwList.Columns.Count - 1
            If dgwList.Columns(i).Visible Then
                dgwSumma.Columns(i).Width = dgwList.Columns(i).Width
                dgwSumma.Columns(i).DefaultCellStyle = style
                width += dgwList.Columns(i).Width
            End If
        Next i
        If Me.Width < width Then
            Me.Panel1.Width = width
        Else
            Me.Panel1.Width = Me.Width
        End If

        style = dgwSumma2.DefaultCellStyle
        style.BackColor = Me.BackColor
        f = New Font(style.Font, Drawing.FontStyle.Bold)
        style.Font = f

        dgwSumma2.Width = dgwList2.Width
        dgwSumma2.RowHeadersWidth = dgwList2.RowHeadersWidth
        If dgwList2.Columns.Count <> dgwSumma2.Columns.Count Then Exit Sub
        For i = 0 To dgwList2.Columns.Count - 1
            dgwSumma2.Columns(i).Width = dgwList2.Columns(i).Width
            dgwSumma2.Columns(i).DefaultCellStyle = style
        Next i

        IsFitdgw = True
        Exit Sub
Err_Hand:
        MsgBox("Ошибка в процедуре frmList.Fitdgw: " & Err.Description)
    End Sub

    Public Sub ReCalcAllSum(ByVal Nom As Integer)
        Dim dgw As DataGridView
        Dim i As Integer

        dgw = GetdgwByNom(Nom)
        'col = GetSumColumnsCollection(Nom)
        For i = 1 To dgw.Columns.Count - 1
            If IsSumColumn(i, Nom) Then ReCalcSum(i, Nom)
        Next i

    End Sub

    Public Function GetdgwByNom(ByVal Nom As Integer) As DataGridView
        If Nom = 1 Then
            Return dgwList
        ElseIf Nom = 2 Then
            Return dgwList2
        End If
        Return Nothing
    End Function

    Public Function GetDataTableByNom(ByVal Nom As Integer) As DataTable
        If Nom = 1 Then
            Return ds
        ElseIf Nom = 2 Then
            Return ds2
        End If
        Return Nothing
    End Function

    Private Sub ReCalcSum(ByVal ColIndex As Integer, ByVal Nom As Integer)
        On Error GoTo Err_h
        Dim i As Integer
        'Dim col As String
        Dim Summa As Double = 0

        If Not IsSumColumn(ColIndex, Nom) Then Exit Sub
        If Nom = 1 Then
            If dgwSumma.Columns.Count <= ColIndex Then Exit Sub
            For i = 0 To dgwList.Rows.Count - 1
                If Not (dgwList.Item(ColIndex, i).Value Is Nothing) Then
                    If dgwList.Item(ColIndex, i).Value.ToString <> "" Then
                        Summa += IIf(IsDBNull(dgwList.Item(ColIndex, i).Value), 0, dgwList.Item(ColIndex, i).Value)
                    End If
                End If
            Next i
            dgwSumma.Item(ColIndex, 0).Value = Format(Summa, "0.00")
        ElseIf Nom = 2 Then
            If dgwSumma2.Columns.Count <= ColIndex Then Exit Sub
            For i = 0 To dgwList2.Rows.Count - 1
                If (dgwList2.Item(ColIndex, i).Value Is Nothing) Then Exit For
                If dgwList2.Item(ColIndex, i).Value.ToString <> "" Then
                    Summa += dgwList2.Item(ColIndex, i).Value
                End If
            Next i
            dgwSumma2.Item(ColIndex, 0).Value = Format(Summa, "0.00")
        End If
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.RecalcSum")

    End Sub

    Private Function GetFieldName(ByVal NomTable As Integer, ByVal NomColumn As Integer) As String
        Dim dt As DataTable
        dt = Populate("select * from " + TableToUpdate(NomTable))
        Using reader As New DataTableReader(dt)
            Return reader.GetName(1)
        End Using

    End Function

    Private Function GetLinkField(ByVal NomTable As Integer) As String
        If mTable(NomTable) = "R_Calc" Then
            Return "ID_Uslugi"
        ElseIf mTable(NomTable) = "R_SF" Then
            Return "ID_VidSF"
        ElseIf mTable(NomTable) = "Q_Uslugi" Then
            Return "ID_Customer"
        Else
            Return GetFieldName(NomTable, 1)
        End If

    End Function

    Public Function GetInsertSQL(ByVal NomTable As Integer) As String
        Dim FieldForLink, sql As String
        Dim i As Integer

        If NomTable = 1 And mTable(1) = "Q_Uslugi" Then
            i = FieldValues("ID_Customer")
            sql = "INSERT INTO R_Uslugi (ID_Customer,Дата_оказания_услуги) VALUES (" + i.ToString + ",?)"
        Else
            FieldForLink = GetLinkField(NomTable)
            If FieldForLink = "" Then FieldForLink = ds.Columns(1).ColumnName
            sql = "INSERT INTO " + TableToUpdate(NomTable) + " (" + FieldForLink.ToString + ") VALUES(" + LinkValue.ToString + ")"
        End If
        Return sql
    End Function

    Public Sub AddCommandParameters(ByVal NomTable As Integer, ByRef Command As OleDbCommand)
        Dim colName As String
        If mTable(NomTable) = "Q_Uslugi" Then
            colName = "Дата_оказания_услуги"
            Command.Parameters.Add( _
                        colName, GetOLEDBType(dgwList.Columns(colName).ValueType.Name), GetOLEDBSize(dgwList.Columns(colName).ValueType.Name))
            Command.Parameters(colName).Value = FieldValues(colName)
        End If
    End Sub

    Public Sub AddRow(ByVal NomTable As Integer, Optional ByVal LinkID As Integer = 0)
        Dim s, FieldForLink As String
        Dim dgw As DataGridView
        Dim sql As String
        Dim dt As DataTable
        Dim dRow As DataRow
        Dim dgwRow As DataGridViewRow
        Dim newID As Integer = 0
        Dim i As Integer
        Dim adap As OleDbDataAdapter
        On Error GoTo Err_h

        If Not IsActivated Then Exit Sub
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If

        dgw = GetdgwByNom(NomTable)
        'adap = New OleDbDataadap()
        dt = GetDataTableByNom(NomTable)
        If NomTable = 1 Then
            'LinkValue = GetLinkValue(NomTable)
            sql = GetInsertSQL(NomTable)
            adap = adapter
            'FieldForLink = GetLinkField(NomTable)
            'If FieldForLink = "" Then FieldForLink = ds.Columns(1).ColumnName
            'sql = "INSERT INTO " + TableToUpdate(NomTable) + " (" + FieldForLink.ToString + ") VALUES(" + LinkValue.ToString + ")"
            adap.InsertCommand = New OleDbCommand(sql, Connection)
            AddCommandParameters(NomTable, adap.InsertCommand)
            adap.InsertCommand.ExecuteNonQuery()
            Dim idCMD As OleDbCommand = New OleDbCommand( _
              "SELECT @@IDENTITY", Connection)

            newID = CInt(idCMD.ExecuteScalar())
        ElseIf NomTable = 2 Then
            If LinkField <> "" Then
                sql = "INSERT INTO " + TableToUpdate(NomTable) + " (" + LinkField + ") VALUES(" + LinkID.ToString + ")"
            Else
                s = GetFieldName(NomTable, 1)
                sql = "INSERT INTO " + TableToUpdate(NomTable) + " (" + s + ") VALUES(Null)"
            End If
            Debug.Print(sql)
            adap = adapter2
            adap.InsertCommand = New OleDbCommand(sql, Connection)
            adap.InsertCommand.ExecuteNonQuery()
            Dim idCMD As OleDbCommand = New OleDbCommand( _
              "SELECT @@IDENTITY", Connection)

            newID = CInt(idCMD.ExecuteScalar())
        End If

        If NomTable = 1 And FieldValues.Count > 0 Then
            FillLinkData(newID)
        End If
        AddingRowBegin(NomTable, newID, Me)
        If mTable(2) = "Q_MedUslug_NoCalc" Then
            dt.Clear()
            If NomTable = 1 Then
                adapter.Fill(dt)
            Else
                adapter2.Fill(dt)
            End If
        Else
            Dim dRows, dRows2 As DataRow()
            'dRows = dt.Select("ID=" + newID.ToString)
            Dim command As New OleDbCommand("SELECT * FROM " + mTable(NomTable) + " WHERE ID=" + newID.ToString, Connection)
            Dim adap2 As New OleDbDataAdapter()
            'If Today.Date > "06/06/2009" Then
            '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
            '    End
            'End If
            adap2.SelectCommand = command
            Dim dt2 As New DataTable
            adap2.Fill(dt2)
            dRows2 = dt2.Select("ID=" + newID.ToString)
            dRow = dt.NewRow
            CopyDataRow(dRows2(0), dRow)
            dt.Rows.Add(dRow)
        End If
        AddingRowEnd(NomTable, newID, Me)

        If NomTable = 2 Then
            'bindingSource2.DataSource = dt
            sql = GetFilterForTable2(LinkID)
            sql = GetFilterForTable2(LinkID)
            Debug.Print(sql)
            bindingSource2.DataSource = ds2
            bindingSource2.Filter = sql
            bindingSource2.ResetBindings(False)
            Debug.Print(bindingSource2.Filter)
            dgw.DataSource = bindingSource2
        End If
        dgw.Refresh()
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.AddRow")
    End Sub

    Public Function IsFilterVisible() As Boolean
        Dim b As Boolean = False
        Dim frm As frmList

        If frmMDI.MdiChildren.Length = 0 Then Exit Function
        frm = frmMDI.ActiveMdiChild
        b = frm.TableName Like "Q_SF_Beznal*"
        Return b
        'Return False
        'For Each frm In frmMDI.MdiChildren
        '    If frm.IsReestr Then
        '        Return True
        '    End If
        'Next frm
        'Return False
    End Function

    Public Sub HideColumns(ByRef dgw As DataGridView)
        Dim i As Integer
        For i = 0 To dgw.Columns.Count - 1
            If (dgw.Columns(i).Name Like "ID*") Then dgw.Columns(i).Visible = False
        Next i
    End Sub

    Private Overloads Sub UpdateCellADO(ByVal NomTable As Integer, ByVal Col As String, ByVal Row As Integer)
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        UpdateCellADO(NomTable, dgw.Columns(Col).Index, Row)

    End Sub

    Private Overloads Sub UpdateCellADO(ByVal NomTable As Integer, ByVal Col As Integer, ByVal Row As Integer)
        Dim dt As DataTable
        Dim dgw As DataGridView
        'Dim adap As New OleDbDataAdapter
        Dim command As New OleDbCommand
        Dim colName As String
        Dim ID_Column As String = "ID"
        Dim sql As String

        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        dgw = GetdgwByNom(NomTable)
        If NomTable = 1 Then
            dt = ds
        Else
            dt = ds2
        End If
        If dgw.Columns(Col).CellType.ToString Like "*ComboBoxCell" Then
            Col = Col - 1
            colName = dgw.Columns(Col).Name
        Else
            colName = dgw.Columns(Col).Name
        End If
        If colName = "ID" Then Exit Sub
        If Not IsColumnInTable(dt, colName) Then Exit Sub
        'If dt.Columns(colName).ReadOnly Or dgw.Columns(colName).Name = "ID" Or IsDBNull(dgw.Item(colName, Row).Value) Then Exit Sub
        If dt.Columns(colName).ReadOnly Or dgw.Columns(colName).Name = "ID" Then Exit Sub
        ID_Column = GetIDColumn(mTable(NomTable), colName)
        If dgw.Item(ID_Column, Row).Value Is Nothing Then Exit Sub
        If dgw.Item(ID_Column, Row).Value.ToString = "" Then Exit Sub

        sql = GetUpdateSql(NomTable, colName, Row)
        If sql = "" Then Exit Sub
        command = New OleDbCommand(sql, Connection)

        command.Parameters.Add( _
            colName, GetOLEDBType(dgw.Columns(colName).ValueType.Name), GetOLEDBSize(dgw.Columns(colName).ValueType.Name))
        'command.Parameters.Add(ID_Column, OleDbType.Integer, 5)
        command.Parameters.Add(ID_Column, GetOLEDBType(dgw.Columns(ID_Column).ValueType.Name), GetOLEDBSize(dgw.Columns(ID_Column).ValueType.Name))
        If dgw.Columns(Col + 1).CellType.ToString Like "*ComboBoxCell" Then
            command.Parameters(colName).Value = dgw.Item(Col + 1, Row).Value
        Else
            command.Parameters(colName).Value = dgw.Item(colName, Row).Value
        End If
        command.Parameters(ID_Column).Value = dgw.Item(ID_Column, Row).Value
        command.ExecuteNonQuery()
        UpdateCellEnd(NomTable, colName, Row)
        dgw.Refresh()

    End Sub

    Private Overloads Sub SetDateValue(ByVal NomTable As Integer, ByVal Col As Integer, ByVal Row As Integer, ByVal cell As DataGridViewCell)
        Dim dt As DataTable
        Dim dgw As DataGridView
        'Dim adap As New OleDbDataAdapter
        Dim command As New OleDbCommand
        Dim colName As String
        Dim ID_Column As String = "ID"
        'Dim Conn As New ADODB.Connection
        Dim sql As String

        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        dgw = GetdgwByNom(NomTable)
        colName = dgw.Columns(Col).Name
        If colName = "ID" Then Exit Sub
        If NomTable = 1 Then
            dt = ds
        Else
            dt = ds2
        End If
        If Not IsColumnInTable(dt, colName) Then Exit Sub
        If dt.Columns(colName).ReadOnly Or dgw.Columns(colName).Name = "ID" Then Exit Sub
        ID_Column = GetIDColumn(mTable(NomTable), colName)
        If dgw.Item(ID_Column, Row).Value Is Nothing Then Exit Sub
        If dgw.Item(ID_Column, Row).Value.ToString = "" Then Exit Sub

        sql = GetUpdateSql(NomTable, colName, Row)
        If sql = "" Then Exit Sub
        command = New OleDbCommand(sql, Connection)

        command.Parameters.Add( _
            colName, GetOLEDBType(dgw.Columns(colName).ValueType.Name), GetOLEDBSize(dgw.Columns(colName).ValueType.Name))
        'command.Parameters.Add(ID_Column, OleDbType.Integer, 5)
        command.Parameters.Add(ID_Column, GetOLEDBType(dgw.Columns(ID_Column).ValueType.Name), GetOLEDBSize(dgw.Columns(ID_Column).ValueType.Name))
        command.Parameters(colName).Value = dgw.Item(colName, Row).Value
        command.Parameters(ID_Column).Value = dgw.Item(ID_Column, Row).Value
        command.ExecuteNonQuery()
        dgw.Refresh()

    End Sub

    Private Overloads Sub UpdateCellEnd(ByVal NomTable As Integer, ByVal ColName As String, ByVal Row As Integer)
        Dim dgw As DataGridView
        Dim sql As String

        dgw = GetdgwByNom(NomTable)
        If ColName.ToLower = "фио_пациента" And (mTable(NomTable) = "Q_DMS" Or mTable(NomTable) = "Q_Stom") Then
            If Not IsDBNull(dgw.Item("ФИО_пациента", Row).Value) Then
                sql = "SELECT * FROM R_FIOLimit WHERE ФИО_пациента='" + dgw.Item("ФИО_пациента", Row).Value.ToString + "'"
                Debug.Print(sql)
                Dim command As New OleDbCommand(sql, Connection)
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If Not reader.Read Then
                    sql = "INSERT INTO R_FIOLimit (ФИО_пациента) VALUES (?)"
                    Dim command2 As New OleDbCommand(sql, Connection)
                    command2.Parameters.Add( _
                                    ColName, GetOLEDBType(dgw.Columns(ColName).ValueType.Name), GetOLEDBSize(dgw.Columns(ColName).ValueType.Name))
                    If dgw.Item(ColName, Row).Value Is Nothing Then Exit Sub
                    If IsDBNull(dgw.Item(ColName, Row).Value) Then Exit Sub
                    command2.Parameters(ColName).Value = dgw.Item(ColName, Row).Value
                    command2.ExecuteNonQuery()
                End If
            End If
            sql = "DELETE FROM R_FIOLimit WHERE (ФИО_пациента Not IN (SELECT DISTINCT ФИО_пациента FROM R_DMS) and ФИО_пациента Not IN (SELECT DISTINCT ФИО_пациента FROM R_Stom)) Or IsNull(ФИО_пациента) Or ФИО_пациента=''"
            Dim command3 As New OleDbCommand(sql, Connection)
            command3.ExecuteNonQuery()
        End If

    End Sub

    Public Function GetUpdateSql(ByVal NomTable As Integer, ByVal ColName As String, Optional ByVal Row As Integer = 0) As String
        Dim sql As String
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        If ColName.ToLower = "код_услуги_по_подразделению" Then
            sql = "UPDATE D_KodMedUslug " & _
                   "SET " + ColName + " = ? " & _
                   "WHERE ID_Uslugi = ? And ID_Year=" + frmMDI.iYear.ToString
        ElseIf ColName.ToLower = "ограничение_по_сумме" Or ColName.ToLower = "ограничение_по_процентам" Then
            sql = "UPDATE R_FIOLimit " & _
                   "SET " + ColName + " = ? " & _
                   "WHERE ФИО_пациента = ?"
        Else
            sql = "UPDATE " + TableToUpdate(NomTable) + " " & _
                   "SET " + ColName + " = ? " & _
                   "WHERE ID = ?"
        End If
        Debug.Print(sql)
        Return sql
    End Function

    Public Sub RefreshAllComboBox(ByVal NomTable As Integer, ByVal Row As Integer)
        Dim dgw As DataGridView
        Dim Col, j As Integer
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        If Row > dgw.Rows.Count - 1 Then
            ErrMess("Неверное значение индекса", "frmList.RefreshComboBox")
            Exit Sub
        End If
        For Col = 0 To dgw.Columns.Count - 1
            If dgw.Columns(Col).CellType.ToString Like "*ComboBoxCell" Then
                If Not IsDBNull(dgw.Item(Col - 1, Row).Value) Then
                    j = dgw.Item(Col - 1, Row).Value
                    IsUpdateRow = True
                    dgw.Item(Col, Row).Value = j
                    IsUpdateRow = False
                End If
            End If
        Next Col
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.RefreshComboBox")
    End Sub

    Public Sub RefreshComboBox(ByVal NomTable As Integer, ByVal Col As Integer, ByVal Row As Integer)
        Dim dgw As DataGridView
        Dim j As Integer
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        If Row > dgw.Rows.Count - 1 Then
            ErrMess("Неверное значение индекса", "frmList.RefreshComboBox")
            Exit Sub
        End If
        If dgw.Columns(Col).CellType.ToString Like "*ComboBoxCell" Then
            If IsDBNull(dgw.Item(Col - 1, Row).Value) Then Exit Sub
            j = dgw.Item(Col - 1, Row).Value
            IsUpdateRow = True
            dgw.Item(Col, Row).Value = j
            IsUpdateRow = False
        End If
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmList.RefreshComboBox")
    End Sub

    Private Sub FillLinkData(ByVal ID As Integer)
        Dim i As Integer
        Dim col As String
        Dim dRow As DataGridViewRow
        Dim dt As DataTable
        On Error GoTo Err_h

        For i = 0 To dgwList.Columns.Count - 1
            col = dgwList.Columns(i).Name
            If IsKeyInCollection(FieldValues, col) Then
                UpdateValueInDgwByID(Me, 1, ID, col, FieldValues(col))
            End If
        Next i
        Exit Sub
Err_h:
        ErrMess(Err.Description, "FillLinkData")
    End Sub

    Private Sub dgwList_DefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgwList.DefaultValuesNeeded
        Dim i As Integer
        Dim col As String

        With e.Row
            For i = 0 To dgwList.Columns.Count - 1
                col = dgwList.Columns(i).Name
                If IsKeyInCollection(FieldValues, col) Then
                    .Cells(col).Value = FieldValues(col)
                End If
            Next i
        End With

    End Sub

    Public Sub RemoveRow(ByVal NomTable As Integer, ByVal Row As Integer)
        Dim id As Integer
        If Me.Visible = False Then Exit Sub
        If Not IsActivated Then Exit Sub
        Dim ds As DataTable
        Dim dRows As DataRow()
        Dim dgw As DataGridView

        ds = GetDataTableByNom(NomTable)
        dgw = GetdgwByNom(NomTable)
        If dgw.Item("ID", Row).Value Is Nothing Then Exit Sub
        If Not IsDBNull(dgw.Item("ID", Row).Value) Then
            id = dgw.Item("ID", Row).Value
            'dgw.Rows.RemoveAt(Row)
            adapter.DeleteCommand = New OleDbCommand("DELETE FROM " + TableToUpdate(NomTable) + " WHERE ID=" + id.ToString, Connection)
            adapter.DeleteCommand.ExecuteNonQuery()
            dRows = ds.Select("ID=" + id.ToString)
            If dRows.Length > 0 Then
                dRows(0).BeginEdit()
                dRows(0).Delete()
                dRows(0).EndEdit()
            End If
            dgw.Refresh()
        Else
            dgw.Rows.RemoveAt(Row)
        End If
        'adapter.Update(bindingSource.DataSource)
        dgw.Refresh()
        ReCalcAllSum(1)
    End Sub

    Private Sub dgwList_HelpRequested(ByVal sender As Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles dgwList.HelpRequested
        If mTable(2) Like "R*" Or mTable(2) Like "Q*" Then
            Help.ShowHelp(Me, frmMDI.hpAdvancedCHM.HelpNamespace, HelpNavigator.Topic, "Reestr.htm")
        Else
            Help.ShowHelp(Me, frmMDI.hpAdvancedCHM.HelpNamespace, HelpNavigator.Topic, "Sprav.htm")
        End If
        'Help.ShowHelp(Me, frmMDI.hpAdvancedCHM.HelpNamespace)
    End Sub

    Private Sub dgwList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgwList.KeyDown
        If e.KeyCode = Keys.Insert And dgwList.AllowUserToAddRows = False Then
            AddRow(1)
        End If
    End Sub

    Private Sub dgwList_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgwList.SelectionChanged
        Dim i, id, row As Integer
        On Error GoTo Err_h

        If dgwList.Visible = False Then Exit Sub
        If Not IsActivated Then Exit Sub
        mnuDel.Visible = (dgwList.SelectedCells.Count > 0)
        mnuCopy.Visible = (dgwList.SelectedCells.Count > 0 And mTable(1) Like "Q_SF*")
        cldr2.Visible = False
        IsAddRowByUser = True
        If dgwList.SelectedCells.Count = 0 Then
            bindingSource2.Filter = "1=0"
            Fitdgw()
            ClearDgwSumma(2)
            IsLinkingWithAnother = False
            'dgwSumma2.Item(4, 0).Value = ""
            Exit Sub
        End If
        If (dgwList.SelectedCells(0).ValueType Is Nothing) Then
            cldr1.Visible = False
        Else
            If dgwList.SelectedCells(0).ValueType.Name <> "DateTime" Then
                cldr1.Visible = False
            End If
        End If
        i = dgwList.SelectedCells(0).RowIndex
        If IsLinkingWithAnother Then
            IsLinkingWithAnother = False
            If IsDBNull(dgwList.Item("ID", i).Value) Then Exit Sub
            LinkWithAnother()
        End If
        If i = SelectedRow(1) And i > 0 Then Exit Sub
        SelectedRow(1) = i
        If IsDBNull(dgwList.Item("ID", i).Value) Then
            IsAddRowByUser = False
            bindingSource2.Filter = "1=0"
            IsAddRowByUser = True
            Fitdgw()
            'dgwSumma2.Item(4, 0).Value = ""
            ClearDgwSumma(2)
            Exit Sub
        End If
        If dgwList.Item("ID", i).Value = 0 Then
            bindingSource2.Filter = "1=0"
        Else
            id = dgwList.Item("ID", i).Value
            IsAddRowByUser = False
            bindingSource2.Filter = GetFilterForTable2(id) 'LinkField & "=" & id.ToString
            IsAddRowByUser = True
            For i = 0 To dgwList2.Columns.Count - 1
                If dgwList2.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                    For row = 0 To dgwList2.Rows.Count - 1
                        RefreshComboBox(2, i, row)
                    Next row
                End If
            Next i
        End If
        HideColumns(dgwList2)
        Fitdgw()
        ReCalcAllSum(2)
        Exit Sub
Err_h:
        ErrMess(Err.Description, "dgwList.SelectionChanged")
    End Sub

    Private Function GetFilterForTable2(ByVal id As Integer) As String
        Dim s As String
        Dim dt As DataTable

        If mTable(2) = "Q_MedUslug_NoCalc" Then
            s = "ID_Year=" + frmMDI.iYear.ToString + " AND "
            s += "ID_Podr=" + id.ToString + " AND "
            s += "ID IN ("
            dt = Populate("select ID_Uslugi from D_MedUslug_Podr where ID_Podr=" + id.ToString)
            Using reader As New DataTableReader(dt)
                Do While reader.Read
                    s += reader("ID_Uslugi").ToString + ","
                Loop
            End Using
            dt.Clear()
            If s = "ID IN (" Then
                s = "1=0"
            Else
                s = Mid(s, 1, Len(s) - 1)
                s += ")"
            End If

        Else
            s = LinkField & "=" & id.ToString
        End If
        Debug.Print(s)
        Return s
    End Function

    Private Sub ClearDgwSumma(ByVal Nom As Integer)
        Dim i As Integer
        On Error GoTo Err_h

        If Nom = 1 Then Exit Sub
        If dgwSumma2.Visible = False Then Exit Sub
        For i = 1 To dgwSumma2.Columns.Count - 1
            dgwSumma2.Item(i, 0).Value = ""
        Next i
        Exit Sub
Err_h:
        ErrMess(Err.Description, "ClearDgwSumma(" + Nom.ToString + ")")
    End Sub

    Public Sub New()

        'Connection = New OleDbConnection(ConnectionString)
        'ADOConn = New ADODB.Connection
        'OpenADOConnection(Me)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ReDim mTable(3)
        ReDim mNoEdit(3)
        ReDim TableToUpdate(3)
        'ReDim IsCellsValueChanged(3)
        ReDim SelectedRow(2)

    End Sub

    Private Sub frmList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not IsAllObligatoryColumnsFill(1) Then
            e.Cancel = True
        End If
        If Not IsAllObligatoryColumnsFill(2) Then
            e.Cancel = True
        End If

        SaveColumnWidthToSettings(1)
        If mTable(2) <> "" Then SaveColumnWidthToSettings(2)

    End Sub

    Private Sub frmList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Integer
        For i = 1 To 2
            If mTable(i) Is Nothing Then mTable(i) = ""
            'If mNoEdit(i) Is Nothing Then mNoEdit(i) = False
        Next i
    End Sub

    Private Sub frmList_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Fitdgw()
    End Sub


    Private Sub cldr1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles cldr1.DateSelected
        cldr1.Visible = False
        If dgwList.SelectedCells.Count = 0 Then
            Exit Sub
        End If
        dgwList.Item(dgwList.SelectedCells(0).ColumnIndex, dgwList.SelectedCells(0).RowIndex).Value = cldr1.SelectionRange.End
    End Sub


    Private Sub mnuDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDel.Click
        Dim row As Integer
        Dim s As String
        Dim Rc As Integer
        Dim Mess As String
        On Error GoTo Err_h

        If dgwList.SelectedCells.Count = 0 Then Exit Sub
        row = dgwList.SelectedCells(0).RowIndex
        If mTable(2) = "" Then
            Mess = "Удалить выделенную запись?"
        Else
            Mess = "Удалить выделенную запись из верхней секции?"
        End If
        If MsgBox(Mess, MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        If mTable(2) = "" Then
            If mTable(1) Like "D_*" Then
                s = LinkedDataInTable(mTable(1), dgwList.Item("ID", row).Value)
                If s = "" Then
                    dgwList.Rows.RemoveAt(row)
                Else
                    MsgBox("Существуют связанные записи в таблице " + s + ". Удаление невозможно.")
                End If
            Else
                RemoveRow(1, row)
            End If
        Else
            RemoveRow(1, row)
        End If
        Exit Sub
Err_h:
        MsgBox("Ошибка при удалении записи.")

    End Sub

    Private Function LinkedDataInTable(ByVal D_Table As String, ByVal ID As String) As String
        Dim Rs, Rs2, Rs3 As OleDbDataReader
        Dim sql As String
        On Error GoTo Err_h

        Rs = readerBySQL("SELECT * FROM M_LinkData WHERE LinkTable='" + D_Table + "'")
        Do While Rs.Read
            sql = "SELECT * FROM " + D_Table + " AS D INNER JOIN " + Rs("Table").ToString + " AS R ON (D.ID=R." + Rs("ID_Column").ToString + ") WHERE D.ID=" + ID.ToString
            Debug.Print(sql)
            Rs2 = readerBySQL(sql)
            If Rs2.Read Then
                Rs3 = readerBySQL("SELECT * FROM M_Tables WHERE Table='" + Rs("Table") + "'")
                If Rs3.Read Then Return Rs3("FormCaption")
            End If
        Loop
        Return ""
        Exit Function
Err_h:
        Return ""
    End Function

    Private Sub dgwList2_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgwList2.CellBeginEdit
        On Error GoTo Err_Hand
        Dim dgw As DataGridView
        Dim row As Integer

        If dgwList.SelectedCells.Count > 0 Then
            row = dgwList.SelectedCells(0).RowIndex
            If dgwList.Rows(row).IsNewRow Then
                e.Cancel = True
                Exit Sub
            End If
        ElseIf dgwList.Rows.Count = 0 Then
            Exit Sub
        Else
            MsgBox("Необходимо выделить запись в верхней секции.")
            e.Cancel = True
            Exit Sub
        End If

        dgw = GetdgwByNom(2)
        If dgw.Columns(CInt(e.ColumnIndex)).ValueType Is Nothing Then
            cldr2.Visible = False
            Exit Sub
        End If
        If dgw.Columns(CInt(e.ColumnIndex)).ValueType.Name <> "DateTime" Then
            cldr2.Visible = False
            Exit Sub
        End If
        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then Exit Sub
        If cldr2.Visible Then Exit Sub
        If dgw.AllowUserToAddRows And e.RowIndex = dgw.Rows.Count - 1 Then Exit Sub
        With cldr2
            .Left = CalendarLeft(2, e.ColumnIndex)
            .Top = CalendarTop(2, e.RowIndex)
            If e.RowIndex < dgw.Rows.Count - 1 And Not (dgw.Item(e.ColumnIndex, e.RowIndex).Value Is System.DBNull.Value) Then
                .SetDate(dgw.Item(e.ColumnIndex, e.RowIndex).Value)
            Else
                .SetDate(Today)
            End If
            .Visible = True
        End With
        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "dgwList2_CellBeginEdit")

    End Sub

    Private Sub dgwList2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList2.CellDoubleClick
        RowDoubleClicked2 = e.RowIndex
        If mTable(2) = "Q_CustomersByPeriod" Then
            frmMDI.ViewForm("Services", Me)
        ElseIf mTable(2) = "Q_MedUslug_NoCalc" Then
            frmMDI.ViewForm("Calculations", Me)
        End If

    End Sub

    Private Sub dgwList2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList2.CellEndEdit
        On Error GoTo Err_Hand
        Dim dgw As DataGridView
        Dim row, id As Integer

        If dgwList.SelectedCells.Count = 0 Then
            Exit Sub
        End If
        dgw = GetdgwByNom(2)
        If dgw.Visible = False Then Exit Sub
        If e.ColumnIndex = 0 Then Exit Sub
        If IsDBNull(dgwList2.Item("ID", e.RowIndex).Value) And newid2 > 0 Then
            dgwList2.Item("ID", e.RowIndex).Value = newid2
            newid2 = 0
        End If
        row = dgwList.SelectedCells(0).RowIndex
        If IsDBNull(dgwList.Item("ID", row).Value) Then
            ErrMess("Значение ID в главной таблице неопределенно.")
            Exit Sub
        End If
        id = dgwList.Item("ID", row).Value
        IsUpdateRow = True
        If IsColumnInDgw(dgwList2, LinkField) Then
            If (dgwList2.Item(LinkField, e.RowIndex).Value Is Nothing) Or IsDBNull(dgwList2.Item(LinkField, e.RowIndex).Value) Then
                dgwList2.Item(LinkField, e.RowIndex).Value = id
                'Updatedgw(2)
            End If
        End If
        IsUpdateRow = False

        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then
            dgw.Item(e.ColumnIndex - 1, e.RowIndex).Value = dgw.Item(e.ColumnIndex, e.RowIndex).Value
        ElseIf dgw.Columns(e.ColumnIndex).ValueType Is Nothing Then
            cldr2.Visible = False
            Exit Sub
        ElseIf dgw.Columns(e.ColumnIndex).ValueType.ToString = "DateTime" Then
            cldr2.Visible = False
        End If
        Exit Sub
Err_Hand:
        ErrMess(Err.Description, "dgwList2_CellEndEdit")
    End Sub

    Private Sub dgwList2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList2.CellValueChanged
        On Error GoTo Err_hand
        Dim dgw As DataGridView
        Dim row, id, i As Integer
        Dim col As String

        'If Not IsActivated Or IsUpdateRow Then Exit Sub
        If Not IsActivated Then Exit Sub
        If dgwList2.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then
            Exit Sub
        End If
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        If e.RowIndex <= -1 Then
            Exit Sub
        End If
        dgw = GetdgwByNom(2)
        If dgwList.SelectedCells.Count = 0 Then Exit Sub
        If e.RowIndex <= -1 Then
            Exit Sub
        End If
        If dgw.Columns(e.ColumnIndex).ReadOnly Then Exit Sub
        UpdateCellADO(2, e.ColumnIndex, e.RowIndex)
        FillRecalcColumns(2, e.RowIndex)
        If IsColumn(e.ColumnIndex, 2, "IsUpdateWhenEnter") Then SaveData()
        FillRecalcColumns(1, dgwList.SelectedCells(0).RowIndex)
        ReCalcSum(e.ColumnIndex, 2)
        Exit Sub
Err_hand:
        MsgBox("Ошибка при изменении значения ячейки " & Err.Description)
    End Sub

    Private Sub dgwList2_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgwList2.ColumnWidthChanged
        On Error GoTo Err_h
        If dgwSumma2.Visible Then
            dgwSumma2.Columns(e.Column.Index).Width = dgwList2.Columns(e.Column.Index).Width
        End If
        Exit Sub
Err_h:
        ErrMess(Err.Description, "dgwList2_ColumnWidthChanged")
    End Sub

    Private Sub dgwList2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgwList2.DataError
        Err.Clear()
    End Sub

    Private Sub dgwList2_DefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgwList2.DefaultValuesNeeded
        Dim row As Integer

        If dgwList.SelectedCells.Count = 0 Then Exit Sub
        row = dgwList.SelectedCells(0).RowIndex
        If Not (LinkField Is Nothing) Then dgwList2.Item(LinkField, e.Row.Index).Value = dgwList.Item("ID", row).Value
    End Sub

    Public Function LinkColumns(ByVal Col As String) As String
        'If Col = "Сумма" And mTable(1) = "R_Calc" Then Return "Сумма"
        'If Col = "Сумма" And (mTable(1) = "Q_Uslugi" Or mTable(1) = "Q_Uslugi" Or mTable(1) = "Q_Uslugi") Then Return "Оплаченная_сумма"
        Dim dRows As DataRow()
        Dim dRow As DataRow

        dRows = dtColAttr.Select("Table='" + mTable(1) + "' AND LinkColumn='" + Col + "'")
        If dRows.Length > 0 Then
            dRow = dRows(0)
            If IsDBNull(dRow("ColumnName")) Then Return ""
            Return dRow("ColumnName")
        End If
        Return ""
    End Function

    Private Sub dgwList2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgwList2.KeyDown
        If e.KeyCode = Keys.Insert And dgwList2.AllowUserToAddRows = False Then
            AddRow(2)
        End If
    End Sub

    Private Sub dgwList2_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgwList2.RowsAdded
        Dim id, row As Integer
        If Me.Visible = False Then Exit Sub
        If Not IsActivated Then Exit Sub
        If Not IsAddRowByUser Then Exit Sub
        If dgwList.Rows.Count = 0 Then Exit Sub
        If dgwList.SelectedCells.Count = 0 And e.RowIndex > 0 Then
            MsgBox("Необходимо выделить запись в верхней секции.")
            Exit Sub
        End If
        If mTable(2) Like "Q*" Then Exit Sub
        If dgwList.SelectedCells.Count = 0 Then Exit Sub
        row = dgwList.SelectedCells(0).RowIndex
        If row <= dgwList.Rows.Count - 1 And IsDBNull(dgwList.Item("ID", row).Value) Then
            ErrMess("Неопределенно значение ID в верхней таблице.", "dgwList2_RowsAdded")
            Exit Sub
        End If
        If e.RowIndex > dgwList2.Rows.Count - 1 Then Exit Sub
        'Updatedgw(2)
        'Dim idCMD As OleDbCommand = New OleDbCommand( _
        '  "SELECT @@IDENTITY", Connection)

        IsAddRowByUser = False
        'newid2 = CInt(idCMD.ExecuteScalar())
        If e.RowIndex > dgwList2.Rows.Count - 1 Then Exit Sub
        'If newid2 > 0 Then dgwList2.Item("ID", e.RowIndex).Value = newid2
        id = dgwList.Item("ID", row).Value
        If Not (LinkField Is Nothing) Then
            UpdateValueInDgw(Me, 2, e.RowIndex, LinkField, id)
            'dgwList2.Item(LinkField, e.RowIndex).Value = id
        End If
        IsAddRowByUser = True
        'SaveData()

    End Sub

    Private Sub dgwList2_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgwList2.RowsRemoved
        If Me.Visible = False Then Exit Sub
        If Not IsActivated Then Exit Sub
        If Not IsAddRowByUser Then Exit Sub
        ReCalcAllSum(2)
    End Sub

    Private Sub dgwSumma2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwSumma2.CellValueChanged
        Dim col, s As String
        Dim row As Integer
        On Error GoTo Err_h

        If dgwList.SelectedCells.Count = 0 Then
            Exit Sub
        Else
            row = dgwList.SelectedCells(0).RowIndex
        End If
        'For Each col In LinkColumns
        col = dgwList2.Columns(e.ColumnIndex).Name
        s = LinkColumns(col)

        IsUpdateRow = True
        dgwList.Item(s, row).Value = dgwSumma2.Item(e.ColumnIndex, 0).Value
        IsUpdateRow = False

        Exit Sub
Err_h:
        ErrMess(Err.Description, "dgwSumma2_CellValueChanged")
    End Sub

    Private Sub cldr2_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles cldr2.DateSelected
        cldr2.Visible = False
        If dgwList2.SelectedCells.Count = 0 Then
            Exit Sub
        End If
        dgwList2.Item(dgwList2.SelectedCells(0).ColumnIndex, dgwList2.SelectedCells(0).RowIndex).Value = cldr2.SelectionRange.End

    End Sub

    Private Sub mnuDel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDel2.Click
        Dim row As Integer

        If dgwList2.SelectedCells.Count = 0 Then Exit Sub
        If MsgBox("Удалить выделенную запись из нижней секции?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        row = dgwList2.SelectedCells(0).RowIndex
        RemoveRow(2, row)
    End Sub

    Private Sub mnuAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdd.Click
        AddRow(1)
    End Sub

    Private Function IsAllObligatoryColumnsFill(ByVal NomTable As String) As Boolean
        Dim dgw As DataGridView
        Dim ColName As String
        Dim row As Integer
        Dim dt As DataTable
        On Error GoTo Err_h

        dgw = GetdgwByNom(NomTable)
        'tab = mTable(NomTable)
        dt = Populate("select * from M_Columns where Table='" + mTable(NomTable) + "' AND IsObligatory=True")
        Using reader As New DataTableReader(dt)
            Do While reader.Read
                ColName = reader("ColumnName")
                For row = 0 To dgw.Rows.Count - 1
                    'If dgw.Item(ColName, row).Value Is Nothing Then Return False
                    If IsDBNull(dgw.Item(ColName, row).Value) Then
                        MsgBox("Не заполнен обязательный столбец '" + ColName.Replace("_", " ") + "' в строке " + (row + 1).ToString)
                        Return False
                    End If
                Next row
            Loop
        End Using
        dt.Clear()
        Return True
Err_h:
        ErrMess(Err.Description, "IsAllObligatoryColumnsFill")
        Return False
    End Function

    Public Sub LinkWithAnother()
        Dim dRows As DataRow()
        Dim id, id2, row As Integer

        If mTable(2) = "" Then Exit Sub
        If dgwList.SelectedCells.Count > 0 Or dgwList2.SelectedCells.Count > 0 Then
            row = dgwList.SelectedCells(0).RowIndex
            If dgwList.Item("ID", row).Value Is Nothing Then Exit Sub
            If IsDBNull(dgwList.Item("ID", row).Value) Then Exit Sub
            id = dgwList.Item("ID", row).Value
            row = dgwList2.SelectedCells(0).RowIndex
            If dgwList2.Item("ID", row).Value Is Nothing Then Exit Sub
            If IsDBNull(dgwList2.Item("ID", row).Value) Then Exit Sub
            id2 = dgwList2.Item("ID", row).Value
            adapter2.UpdateCommand = New OleDbCommand("UPDATE " + TableToUpdate(2) + " SET " + LinkField + "=" + id.ToString + " WHERE ID=" + id2.ToString, Connection)
            adapter2.UpdateCommand.ExecuteNonQuery()
            dRows = ds2.Select("ID=" + id2.ToString)
            If dRows.Length > 0 Then
                dRows(0).BeginEdit()
                dRows(0)(LinkField) = id
                dRows(0).EndEdit()
                IsLinkingWithAnother = False
            End If
        End If
    End Sub

    Public Sub GetLinkColumns(ByVal NomTable As Integer)
        If dtLinkCols.Columns.Count = 0 Then
            dtLinkCols = Populate("select * from M_LinkData")
        End If

    End Sub

    Public Sub GetColumnsAttr(ByVal NomTable As Integer)
        If dtColAttr.Columns.Count = 0 Then
            dtColAttr = Populate("select * from M_Columns")
        End If

    End Sub

    Public Sub GetTablesAttr()
        If dtTablesAttr.Columns.Count = 0 Then
            dtTablesAttr = Populate("select * from M_Tables where Table='" + mTable(1) + "' OR Table='" + mTable(2) + "'")
        End If

    End Sub

    Private Sub mnuAdd2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdd2.Click
        Dim ID1 As Integer
        Dim dgw As DataGridView
        Dim Cell As DataGridViewCell

        dgw = GetdgwByNom(1)
        If dgw.SelectedCells.Count = 0 Then Exit Sub
        Cell = dgw.SelectedCells(0)
        If Not IsCellValueDefined(1, "ID", Cell.RowIndex) Then Exit Sub
        Cell = dgw.Item("ID", Cell.RowIndex)
        ID1 = Cell.Value
        AddRow(2, ID1)
        'adapter2.Update(ds2)
        'AddingRowEnd(2, , Me)
    End Sub

    Public Overloads Function IsCellValueDefined(ByVal NomTable As Integer, ByVal Col As String, ByVal Row As Integer) As Boolean
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        If dgw.Item(Col, Row).Value Is Nothing Then Return False
        If IsDBNull(dgw.Item(Col, Row).Value) Then Return False
        Return True
    End Function

    Public Overloads Function IsCellValueDefined(ByVal NomTable As Integer, ByVal Col As Integer, ByVal Row As Integer) As Boolean
        Dim dgw As DataGridView

        dgw = GetdgwByNom(NomTable)
        Return IsCellValueDefined(NomTable, dgw.Columns(Col).Name, Row)
    End Function

    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        CopySelSF()
    End Sub

    Public Function CopySelSF() As Boolean
        Dim SelRow, ID_Sf, newID, i, Row As Integer
        Dim sql As String
        Dim dt As DataTable
        On Error GoTo Err_h

        If dgwList.SelectedCells.Count = 0 Then
            MsgBox("Необходимо выбрать счет-фактуру.")
            Return False
        End If
        IsCopy = True
        SelRow = dgwList.SelectedCells(0).RowIndex
        ID_Sf = dgwList.Item("ID", SelRow).Value
        sql = "INSERT INTO R_SF (ID_VidSF, ID_StrahComp,ID_VidMedStrah,ID_Customer,ID_Podr,Номер,Дата) SELECT ID_VidSF, ID_StrahComp,ID_VidMedStrah,ID_Customer,ID_Podr,Номер,Дата FROM R_SF WHERE ID=" + ID_Sf.ToString
        Dim Command As New OleDbCommand(sql, Connection)
        Command.ExecuteNonQuery()
        Dim idCMD As OleDbCommand = New OleDbCommand( _
                      "SELECT @@IDENTITY", Connection)

        newID = CInt(idCMD.ExecuteScalar())
        Select Case ID_VidSF
            Case 1 'военнослужащие-амбулаторно
                sql = "INSERT INTO R_SoldAmb (ID_SF, ID_Podr,ID_Uslugi,Сумма) SELECT " + newID.ToString + ", ID_Podr,ID_Uslugi,Сумма FROM R_SoldAmb WHERE ID_SF=" + ID_Sf.ToString
            Case 2 'военнослужащие по стационарам
                sql = "INSERT INTO R_SoldSt (ID_SF, ID_Podr,ID_Uslugi,Дата_начала_лечения, Дата_окончания_лечения, ФИО_Пациента, Диагноз_МКБ, МЭС,Количество_оказанной_услуги) SELECT " + newID.ToString + ", ID_Podr,ID_Uslugi,Дата_начала_лечения, Дата_окончания_лечения, ФИО_Пациента, Диагноз_МКБ, МЭС,Количество_оказанной_услуги FROM R_SoldSt WHERE ID_SF=" + ID_Sf.ToString
            Case 3 'ДМС
                sql = "INSERT INTO R_DMS (ID_SF, ID_Podr,ID_Uslugi,Дата_посещения, Номер_полиса,ФИО_Пациента, Диагноз_МКБ, Количество_оказанной_услуги) SELECT " + newID.ToString + ", ID_Podr,ID_Uslugi,Дата_посещения, Номер_полиса,ФИО_Пациента, Диагноз_МКБ,Количество_оказанной_услуги FROM R_DMS WHERE ID_SF=" + ID_Sf.ToString
            Case 4 'безнал-факт
                sql = "INSERT INTO R_SpecSF (ID_SF, ID_Podr,ID_Uslugi,Количество_оказанной_услуги) SELECT " + newID.ToString + ", ID_Podr,ID_Uslugi,Количество_оказанной_услуги FROM R_SpecSF WHERE ID_SF=" + ID_Sf.ToString
            Case 5 'безнал-проект
                sql = "INSERT INTO R_SpecSF (ID_SF, ID_Podr,ID_Uslugi,Количество_оказанной_услуги) SELECT " + newID.ToString + ", ID_Podr, ID_Uslugi,Количество_оказанной_услуги FROM R_SpecSF WHERE ID_SF=" + ID_Sf.ToString
            Case 6 'стоматология
                sql = "INSERT INTO R_Stom (ID_SF, ID_Uslugi,Количество_оказанной_услуги) SELECT " + newID.ToString + ", ID_Uslugi,Количество_оказанной_услуги FROM R_SpecSF WHERE ID_SF=" + ID_Sf.ToString
        End Select
        Debug.Print(sql)
        Dim Command2 As New OleDbCommand(sql, Connection)
        Command2.ExecuteNonQuery()
        dt = GetDataTableByNom(1)
        dt.Clear()
        adapter.Fill(dt)
        For i = 0 To dgwList.Columns.Count - 1
            If dgwList.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                For Row = 0 To dgwList.Rows.Count - 1
                    RefreshComboBox(1, i, Row)
                Next Row
            End If
        Next i
        IsCopy = False

        Exit Function
Err_h:
        IsCopy = False
        ErrMess("Произошла ошибка при копировании счет-фактуры: " + Err.Description)

    End Function

    Private Sub dgwList_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgwList.Sorted
        Dim i, Row As Integer
        For i = 0 To dgwList.Columns.Count - 1
            If dgwList.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                For Row = 0 To dgwList.Rows.Count - 1
                    RefreshComboBox(1, i, Row)
                Next Row
            End If
        Next i

    End Sub

    Private Sub dgwList2_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgwList2.Sorted
        Dim i, Row As Integer
        For i = 0 To dgwList2.Columns.Count - 1
            If dgwList2.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
                For Row = 0 To dgwList2.Rows.Count - 1
                    RefreshComboBox(2, i, Row)
                Next Row
            End If
        Next i

    End Sub

End Class

