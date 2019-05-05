Module Archive
    'Private Sub cldr2_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs)
    '    cldr2.Visible = False
    '    If dgwList2.SelectedCells.Count = 0 Then
    '        Exit Sub
    '    End If
    '    dgwList2.Item(dgwList2.SelectedCells(0).ColumnIndex, dgwList2.SelectedCells(0).RowIndex).Value = cldr2.SelectionRange.End

    'End Sub

    'Private Sub dgwList2_RowsRemoved1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs)
    '    If Me.Visible = False Then Exit Sub
    '    If Not IsActivated Then Exit Sub
    '    If Not IsAddRowByUser Then Exit Sub
    '    ReCalcAllSum(2)
    'End Sub

    'Private Sub mnuAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim frm As frmList
    '    frm = Me.ActiveMdiChild
    '    frm.AddRow(1)
    'End Sub

    'Private Sub mnuAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdd.Click
    '    Dim frm As frmList
    '    frm = frmMDI.ActiveMdiChild
    '    frm.AddRow(1)
    'End Sub

    '    Public Sub UpdateDataTable(ByVal NomTable As Integer)
    '        Dim commandBuilder As OleDbCommandBuilder
    '        Dim dgw As DataGridView
    '        Dim i As Integer
    '        Dim col, s As String

    '        On Error GoTo Err_h
    '        'AddHandler adapter.RowUpdated, _
    '        'New OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

    '        If Today.Date > "02/02/2009" Then
    '            MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
    '            End
    '        End If
    '        If NomTable = 1 Then
    '            i = ds.Rows.Count
    '            Debug.Print("До " + i.ToString)
    '            ds.Clear()
    '            bindingSource.DataSource = Nothing
    '            ds = Nothing
    '            ds = New DataTable
    '            adapter = Nothing
    '            commandBuilder = New OleDbCommandBuilder(adapter)
    '            adapter = New OleDb.OleDbDataAdapter("select * from " + mTable(NomTable) + " where " + Filter, Connection)
    '            adapter.Fill(ds)
    '            Debug.Print(adapter.SelectCommand.CommandText)
    '            i = ds.Rows.Count
    '            Debug.Print("После " + i.ToString)
    '            bindingSource.DataSource = ds
    '            bindingSource.ResetBindings(False)
    '            dgwList.DataSource = bindingSource
    '            dgwList.Refresh()
    '        Else
    '            ds2.Clear()
    '            adapter2 = Nothing
    '            adapter2 = New OleDb.OleDbDataAdapter("select * from " + mTable(NomTable) + " where " + Filter, Connection)
    '            commandBuilder = New OleDbCommandBuilder(adapter2)
    '            adapter2.Fill(ds2)
    '        End If
    '        dgw = GetdgwByNom(NomTable)
    '        'dgw.Columns.Add("Dummy", "")
    '        dgw.Columns(dgwList.Columns.Count - 1).ReadOnly = True
    '        SetNoEdit(NomTable, mNoEdit(NomTable))
    '        If dgw.Columns.Count > 5 Then dgw.Columns(dgw.Columns.Count - 1).FillWeight = 0.05
    '        'If dgw.Columns.Count > 5 Then dgw.Columns(dgw.Columns.Count - 1).MinimumWidth = 10
    '        For i = 1 To dgw.Columns.Count - 1
    '            If dgw.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
    '                RefreshLinkColumnData(i - 1, i, NomTable)
    '            End If
    '        Next i
    '        For i = 1 To dgw.Columns.Count - 1
    '            If IsColumnRecalc(i, NomTable) Then FillLinkColumn(NomTable, dgw.Columns(i).Name)
    '        Next i
    '        Fitdgw()
    '        dgw.Refresh()
    '        Exit Sub
    'Err_h:
    '        ErrMess(Err.Description, "frmList.UpdateDataTable")

    '    End Sub

    'Private Sub mnuAdd2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdd2.Click
    '    Dim frm As frmList
    '    Dim id, i As Integer

    '    frm = Me
    '    With frm
    '        If .dgwList.SelectedCells.Count = 0 Then Exit Sub
    '        'Row=frm.dgwList.Item(
    '        i = .dgwList.SelectedCells(0).RowIndex
    '        If IsDBNull(.dgwList.Item("ID", i).Value) Then Exit Sub
    '        If .dgwList.Item("ID", i).Value = 0 Then Exit Sub
    '        id = .dgwList.Item("ID", i).Value
    '        .AddRow(2, id)
    '        For i = 0 To .dgwList2.Columns.Count - 1
    '            If .dgwList2.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
    '                .RefreshLinkColumnData(i - 1, i, 2)
    '            End If
    '        Next i
    '        .HideColumns(.dgwList2)
    '        .Fitdgw()
    '        .ReCalcAllSum(2)
    '    End With
    'End Sub

    'Private Sub dgwList2_RowsAdded1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
    '    Dim id, row As Integer
    '    If Me.Visible = False Then Exit Sub
    '    If Not IsActivated Then Exit Sub
    '    If Not IsAddRowByUser Then Exit Sub
    '    If dgwList.SelectedCells.Count = 0 Then Exit Sub
    '    If mTable(2) Like "Q*" Then Exit Sub
    '    row = dgwList.SelectedCells(0).RowIndex
    '    id = dgwList.Item("ID", row).Value
    '    IsUpdateRow = True
    '    dgwList2.Item(LinkField, e.RowIndex).Value = id
    '    IsUpdateRow = False
    '    IsAddRowByUser = False
    '    adapter2.Update(ds2)
    '    IsAddRowByUser = True

    'End Sub
    '    Private Sub dgwList2_CellBeginEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs)
    '        On Error GoTo Err_Hand
    '        Dim dgw As DataGridView

    '        dgw = GetdgwByNom(2)
    '        If dgw.Columns(CInt(e.ColumnIndex)).ValueType Is Nothing Then
    '            cldr2.Visible = False
    '            Exit Sub
    '        End If
    '        If dgw.Columns(CInt(e.ColumnIndex)).ValueType.Name <> "DateTime" Then
    '            cldr2.Visible = False
    '            Exit Sub
    '        End If
    '        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then Exit Sub
    '        If cldr2.Visible Then Exit Sub
    '        If dgw.AllowUserToAddRows And e.RowIndex = dgw.Rows.Count - 1 Then Exit Sub
    '        With cldr2
    '            .Left = CalendarLeft(2, e.ColumnIndex)
    '            .Top = CalendarTop(2, e.RowIndex)
    '            If e.RowIndex < dgw.Rows.Count - 1 And Not (dgw.Item(e.ColumnIndex, e.RowIndex).Value Is System.DBNull.Value) Then
    '                .SetDate(dgw.Item(e.ColumnIndex, e.RowIndex).Value)
    '            Else
    '                .SetDate(Today)
    '            End If
    '            .Visible = True
    '        End With
    '        Exit Sub
    'Err_Hand:
    '        ErrMess(Err.Description, "dgwList2_CellBeginEdit")

    '    End Sub

    'Private Sub dgwList2_CellDoubleClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '    If mTable(2) = "Q_CustomersByPeriod" Then
    '        frmMDI.ViewForm("Services", Me)
    '    End If
    'End Sub

    '    Private Sub dgwList2_CellEndEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '        On Error GoTo Err_Hand
    '        Dim dgw As DataGridView

    '        dgw = GetdgwByNom(1)
    '        If dgw.Visible = False Then Exit Sub
    '        If e.ColumnIndex = 0 Then Exit Sub
    '        If dgw.Columns(e.ColumnIndex).CellType.ToString Like "*ComboBoxCell" Then
    '            dgw.Item(e.ColumnIndex - 1, e.RowIndex).Value = dgw.Item(e.ColumnIndex, e.RowIndex).Value
    '        ElseIf dgwList.Columns(e.ColumnIndex).ValueType Is Nothing Then
    '            cldr2.Visible = False
    '            Exit Sub
    '        ElseIf dgwList.Columns(e.ColumnIndex).ValueType.ToString = "DateTime" Then
    '            cldr2.Visible = False
    '        End If
    '        Exit Sub
    'Err_Hand:
    '        ErrMess(Err.Description, "dgwList2_CellEndEdit")
    '    End Sub

    'Private Sub dgwList2_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList2.CellLeave
    '    If cldr2.Visible Then cldr2.Visible = False
    'End Sub

    '    Private Sub dgwList_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgwList.CellValueChanged
    '        Dim i As Integer
    '        Dim col As String
    '        On Error GoTo Err_Hand
    '        If Not IsActivated Then Exit Sub
    '        If IsAddRowByUser Then Exit Sub
    '        If Today.Date > "02/02/2009" Then
    '            MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
    '            End
    '        End If
    '        If IsUpdateRow Then
    '            ReCalcSum(e.ColumnIndex, 1)
    '            Exit Sub
    '        End If
    '        If e.RowIndex <= -1 Then
    '            Exit Sub
    '        End If
    '        'If IsUpCellsT1 Then UpdateCellADO(1, e.ColumnIndex, e.RowIndex)
    '        For i = 1 To dgwList.Columns.Count - 1
    '            col = dgwList.Columns(i).Name
    '            If IsColumnFireFillLinks(i, 1) Then
    '                UpdateCellADO(1, i, e.RowIndex)
    '                FillLinkColumn(1, col, , e.RowIndex)
    '            End If
    '            If IsColumnRecalc(i, 1) Then
    '                IsUpdateRow = True
    '                dgwList.EditMode = DataGridViewEditMode.EditProgrammatically
    '                col = dgwList.Columns(i).Name
    '                FillLinkColumn(1, col, , e.RowIndex)
    '                dgwList.EditMode = DataGridViewEditMode.EditOnKeystroke
    '                IsUpdateRow = False
    '            End If
    '        Next i
    '        'ReCalcAllSum(1)
    '        Exit Sub
    'Err_Hand:
    '        ErrMess(Err.Description, "dgwList_CellValueChanged")
    '    End Sub

    '    Private Sub dgwList2_CellValueChanged1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '        On Error GoTo Err_hand
    '        Dim dgw As DataGridView
    '        Dim row, id, i As Integer
    '        Dim col As String

    '        If Not IsActivated Or IsUpdateRow Then Exit Sub
    '        If Today.Date > "02/02/2009" Then
    '            MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
    '            End
    '        End If
    '        If e.RowIndex <= -1 Then
    '            Exit Sub
    '        End If
    '        dgw = GetdgwByNom(2)
    '        If dgwList.SelectedCells.Count = 0 Then Exit Sub
    '        'For i = 1 To dgw.Columns.Count - 1
    '        i = e.ColumnIndex
    '        col = dgw.Columns(i).Name
    '        If IsSumColumn(e.ColumnIndex, 2) Then
    '            ReCalcSum(e.ColumnIndex, 2)
    '        End If
    '        If IsColumnDataUpdateWhenEnter(i, 2) Then
    '            UpdateCellADO(2, i, e.RowIndex)
    '        End If
    '        If IsColumnFireFillLinks(i, 2) Then
    '            FillLinkColumn(2, col, , e.RowIndex)
    '        End If
    '        If IsColumnRecalc(i, 2) Then
    '            IsUpdateRow = True
    '            dgw.EditMode = DataGridViewEditMode.EditProgrammatically
    '            col = dgw.Columns(i).Name
    '            FillLinkColumn(2, col, , e.RowIndex)
    '            dgw.EditMode = DataGridViewEditMode.EditOnKeystroke
    '            IsUpdateRow = False
    '        End If
    '        'Next i
    '        row = dgwList.SelectedCells(0).RowIndex
    '        id = dgwList.Item("ID", row).Value
    '        IsUpdateRow = True
    '        If IsDBNull(dgwList2.Item(LinkField, e.RowIndex).Value) Then dgwList2.Item(LinkField, e.RowIndex).Value = id
    '        IsUpdateRow = False
    '        Exit Sub
    'Err_hand:
    '        MsgBox("Ошибка при изменении значения ячейкм " & Err.Description)
    '    End Sub

    '    Private Sub dgwList2_ColumnWidthChanged1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
    '        If dgwSumma2.Visible Then
    '            dgwSumma2.Columns(e.Column.Index).Width = dgwList2.Columns(e.Column.Index).Width
    '        End If
    '    End Sub

    '    Private Sub dgwList2_DataError1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
    '        Err.Clear()
    '    End Sub

    '    Private Sub dgwSumma2_CellValueChanged1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '        Dim col, s As String
    '        Dim row As Integer

    '        If dgwList.SelectedCells.Count = 0 Then
    '            Exit Sub
    '        Else
    '            row = dgwList.SelectedCells(0).RowIndex
    '        End If
    '        'For Each col In LinkColumns
    '        col = dgwList2.Columns(e.ColumnIndex).Name
    '        On Error GoTo Err_h
    '        s = LinkColumns(col)

    '        IsUpdateRow = True
    '        dgwList.Item(s, row).Value = dgwSumma2.Item(e.ColumnIndex, 0).Value
    '        IsUpdateRow = False

    '        Exit Sub
    '        Exit Sub
    'Err_h:
    '        ErrMess(Err.Description)
    '    End Sub
    '   Private Sub dgwList_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgwList.RowsAdded

    'If Me.Visible = False Or IsUpdateRow Then Exit Sub
    'If Not IsActivated Then Exit Sub

    'adapter.InsertCommand = New OleDbCommand("INSERT INTO " + TableToUpdate(1) + " (" + dgwList.Columns(1).Name + ") VALUES(0)", Connection)
    'adapter.InsertCommand.ExecuteNonQuery()
    'Dim idCMD As OleDbCommand = New OleDbCommand( _
    '  "SELECT @@IDENTITY", Connection)

    'newid = CInt(idCMD.ExecuteScalar())
    'dgwList.Item("ID", e.RowIndex).Value = newid
    'If mTable(1) = "R_Calc" Then AddStatCalc(newid)
    'ReCalcAllSum(1)
    'IsAddRow = True
    '    End Sub
    'Private Sub UpdateCell(ByVal NomTable As Integer, ByVal Col As Integer, ByVal Row As Integer)
    '    Dim dt As DataTable
    '    Dim dgw As DataGridView
    '    Dim adap As New OleDbDataAdapter
    '    Dim sql, colName As String
    '    Dim Param As OleDbParameter

    '    UpdateCellADO(NomTable, Col, Row)
    '    Exit Sub
    '    If Not IsUpCellsT1 And NomTable = 1 Then Exit Sub
    '    dgw = GetdgwByNom(NomTable)
    '    colName = dgw.Columns(Col).Name
    '    If colName = "ID" Then Exit Sub
    '    If NomTable = 1 Then
    '        dt = ds
    '    Else
    '        dt = ds2
    '    End If
    '    If Not IsColumnInTable(dt, colName) Then Exit Sub
    '    If dt.Columns(colName).ReadOnly Or dgw.Columns(colName).Name = "ID" Or IsDBNull(dgw.Item(colName, Row).Value) Then Exit Sub
    '    If dgw.Item("ID", Row).Value.ToString = "" Then Exit Sub
    '    If dgw.Columns(Col).CellType.ToString Like "*ComboBoxCell" Then
    '        sql = "UPDATE " + TableToUpdate(NomTable).ToString + " SET " + dgw.Columns(Col - 1).Name.ToString + "=? WHERE ID=" + dgw.Item("ID", Row).Value.ToString
    '        adap.UpdateCommand = New OleDbCommand(sql, Connection)
    '        Param = adap.UpdateCommand.Parameters.Add("Param", OleDbType.BigInt)
    '        Param.Value = dgw.Item(Col - 1, Row).Value
    '    Else
    '        sql = "UPDATE " + TableToUpdate(NomTable).ToString + " SET " + dgw.Columns(Col).Name.ToString + "=? WHERE ID=" + dgw.Item("ID", Row).Value.ToString
    '        adap.UpdateCommand = New OleDbCommand(sql, Connection)
    '        Param = adap.UpdateCommand.Parameters.Add("Param", GetOLEDBType(dgw.Columns(Col).ValueType.Name))
    '        Param.Value = dgw.Item(Col, Row).Value
    '    End If
    '    adap.UpdateCommand.ExecuteNonQuery()


    'End Sub
    '    Public Sub Updatedgw(ByVal NomTable As Integer)
    '        Dim i As Integer
    '        Dim dgw As DataGridView
    '        Dim commandBuilder As OleDbCommandBuilder
    '        On Error GoTo Err_h

    '        If Not IsAddRowByUser Then Exit Sub
    '        IsAddRowByUser = False
    '        dgw = GetdgwByNom(NomTable)
    '        For i = 1 To dgw.Columns.Count - 1
    '            If dgw.Columns(i).CellType.ToString Like "*ComboBoxCell" Then
    '                RefreshLinkColumnData(i - 1, i, NomTable)
    '            End If
    '        Next i
    '        If NomTable = 1 Then
    '            'adapter = Nothing
    '            dgwList.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '            bindingSource.EndEdit()
    '            adapter = New OleDb.OleDbDataAdapter("select * from " + mTable(NomTable) + " where " + Filter, Connection)
    '            commandBuilder = New OleDbCommandBuilder(adapter)
    '            adapter.Update(bindingSource.DataSource)
    '            'i = ds.Rows.Count
    '            'Debug.Print("После " + i.ToString)
    '            'bindingSource.ResetBindings(False)
    '            'dgwList.DataSource = bindingSource
    '            dgwList.Refresh()
    '        ElseIf NomTable = 2 Then
    '            dgwList2.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '            bindingSource2.EndEdit()
    '            adapter2.Update(bindingSource2.DataSource)
    '        End If
    '        For i = 1 To dgw.Columns.Count - 1
    '            If IsColumnRecalc(i, NomTable) Then FillLinkColumn(NomTable, dgw.Columns(i).Name)
    '        Next i
    '        IsAddRowByUser = True
    '        Exit Sub
    'Err_h:
    '        IsAddRowByUser = True
    '        ErrMess(Err.Description, "frmList.Updatedgw")
    '        Err.Clear()
    '    End Sub
    '    Public Sub FillLinkColumn(ByVal NomTable As Integer, Optional ByVal Col As String = "", Optional ByVal LinkTable As String = "", Optional ByVal Row As Integer = 0)
    '        Dim Rs As New ADODB.Recordset
    '        Dim q As String
    '        Dim i, colIndex As Integer
    '        Dim dgw As DataGridView
    '        Dim s As Single
    '        On Error GoTo Err_Hand

    '        If Today.Date > "02/02/2009" Then
    '            MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
    '            End
    '        End If
    '        dgw = GetdgwByNom(NomTable)
    '        If LinkTable = "" Then
    '            q = mTable(NomTable).ToString.Replace("R_", "Q_")
    '        Else
    '            q = LinkTable
    '        End If
    '        If Row >= dgw.Rows.Count - 1 Then Exit Sub
    '        Rs.Open("SELECT * FROM " & q, ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
    '        'dt = Populate("SELECT * FROM " & q)
    '        'ColumnInLinkTable = Col
    '        If Row = 0 Then
    '            For Row = 0 To dgw.Rows.Count - 1
    '                If IsDBNull(dgw.Item("ID", Row).Value) Or dgw.Item("ID", Row).Value Is Nothing Then Exit For
    '                'dRows = dt.Select("ID=" + dgw.Item("ID", Row).Value.ToString)
    '                Rs.Filter = "ID=" + dgw.Item("ID", Row).Value.ToString
    '                If Not Rs.EOF Then
    '                    IsUpdateRow = True
    '                    'dgw.Item(Col, Row).Value = dRow(Col)
    '                    s = Rs(Col).Value
    '                    dgw.Item(Col, Row).Value = s
    '                    'UpdateCellADO(NomTable, Col, Row)
    '                    IsUpdateRow = False
    '                End If
    '            Next Row
    '        Else
    '            If IsDBNull(dgw.Item("ID", Row).Value) Or dgw.Item("ID", Row).Value Is Nothing Then Exit Sub
    '            Rs.Filter = "ID=" + dgw.Item("ID", Row).Value.ToString
    '            If Not Rs.EOF Then
    '                IsUpdateRow = True
    '                'dgw.Item(Col, Row).Value = dRow(Col)
    '                s = Rs(Col).Value
    '                dgw.Item(Col, Row).Value = s
    '                'UpdateCellADO(NomTable, Col, Row)
    '                IsUpdateRow = False
    '            End If
    '        End If
    '        If Col = "" Then
    '            ReCalcAllSum(NomTable)
    '        Else
    '            colIndex = dgw.Columns(Col).Index
    '            ReCalcSum(colIndex, NomTable)
    '        End If

    '        Exit Sub
    'Err_Hand:
    '        MsgBox("Ошибка в процедуре frmList.FillLinkColumn: " & Err.Description)

    '    End Sub


    '    Public Sub FillLinkColumn(ByVal NomTable As Integer, Optional ByVal Col As String = "", Optional ByVal LinkTable As String = "", Optional ByVal Row As Integer = 0)
    '        Dim dt As DataTable
    '        Dim dRows As DataRow()
    '        Dim dRow As Data.DataRow
    '        Dim q As String
    '        Dim i As Integer
    '        Dim dgw As DataGridView
    '        Dim s As Single
    '        On Error GoTo Err_Hand

    '        If Today.Date > "02/02/2009" Then
    '            MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
    '            End
    '        End If
    '        dgw = GetdgwByNom(NomTable)
    '        If LinkTable = "" Then
    '            q = mTable(NomTable).ToString.Replace("R_", "Q_")
    '        Else
    '            q = LinkTable
    '        End If
    '        dt = Populate("SELECT * FROM " & q)
    '        'ColumnInLinkTable = Col
    '        If Row = 0 Then
    '            For Row = 0 To dgw.Rows.Count - 1
    '                If IsDBNull(dgw.Item("ID", Row).Value) Or dgw.Item("ID", Row).Value Is Nothing Then Exit For
    '                dRows = dt.Select("ID=" + dgw.Item("ID", Row).Value.ToString)
    '                If Not (dRows Is Nothing) Then
    '                    If dRows.Length > 0 Then
    '                        'dRow = dRows(0)
    '                        dRow = dt.Select("ID=" + dgw.Item("ID", Row).Value.ToString)(0)
    '                        IsUpdateRow = True
    '                        'dgw.Item(Col, Row).Value = dRow(Col)
    '                        s = dRow(Col)
    '                        dgw.Item(Col, Row).Value = s
    '                        Debug.Print("Значение=" + s.ToString)
    '                        IsUpdateRow = False
    '                    End If
    '                End If
    '            Next Row
    '        Else
    '            If IsDBNull(dgw.Item("ID", Row).Value) Then Exit Sub
    '            dRows = dt.Select("ID=" + dgw.Item("ID", Row).Value.ToString)
    '            If dRows Is Nothing Then Exit Sub
    '            If dRows.Length > 0 Then
    '                dRow = dt.Select("ID=" + dgw.Item("ID", Row).Value.ToString)(0)
    '                IsUpdateRow = True
    '                's = dRow(Col)
    '                'dgw.Item(ColumnInLinkTable, Row).Value = s
    '                dgw.Item(Col, Row).Value = dRow(Col)
    '                IsUpdateRow = False
    '            End If
    '        End If
    '        ReCalcAllSum(NomTable)
    '        Exit Sub
    'Err_Hand:
    '        MsgBox("Ошибка в процедуре frmList.FillLinkColumns: " & Err.Description)

    '    End Sub
    'Private Function IsSumColumn(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
    '    On Error Resume Next
    '    Dim s As String

    '    Err.Clear()
    '    If NomTable = 1 Then
    '        s = SumColumns(dgwList.Columns(ColIndex).Name)
    '        Return Err.Number = 0
    '    ElseIf NomTable = 2 Then
    '        s = SumColumns2(dgwList2.Columns(ColIndex).Name)
    '        Return Err.Number = 0
    '    End If
    'End Function

    'Private Sub GetReCalcColumns(ByVal tab As String, ByRef Col As Collection)
    '    Dim dt As DataTable
    '    If Col.Count > 0 Then Col.Clear()
    '    'dt = Populate("select * from M_ReCalcColumns where Table='" + tab + "'")
    '    dt = Populate("select * from M_LinkData where (LinkColumnType=2 Or LinkColumnType=3) AND Table='" + tab + "' ORDER BY ID")
    '    Using reader As New DataTableReader(dt)
    '        Do While reader.Read
    '            Try
    '                Col.Add(reader("Column").ToString, reader("Column").ToString)
    '            Catch ex As Exception
    '                Err.Clear()
    '            End Try
    '        Loop
    '    End Using

    'End Sub

    'Private Sub GetSumColumns(ByVal tab As String, ByRef Col As Collection)
    '    Dim dt As DataTable
    '    Dim s As String

    '    If Col.Count > 0 Then Col.Clear()
    '    dt = Populate("select * from M_SumColumns where Table='" + tab + "'")
    '    Try
    '        Using reader As New DataTableReader(dt)
    '            Do While reader.Read
    '                s = reader("Column").ToString()
    '                Col.Add(s, s)
    '            Loop
    '        End Using
    '    Catch ex As Exception
    '        MsgBox("Не удалось получить вычисляемые столбцы. Ошибка:" + ex.Message)
    '    End Try

    'End Sub

    'Private Sub GetLinkColumns(ByVal tab As String)
    '    Dim dt As DataTable
    '    Dim Col As Collection

    '    Col = LinkColumns
    '    If Col.Count > 0 Then Col.Clear()
    '    dt = Populate("select * from M_LinkData where LinkColumnType=3 and Table='" + tab + "'")
    '    Try
    '        Using reader As New DataTableReader(dt)
    '            Do While reader.Read
    '                Col.Add(reader("Column").ToString, reader("ColumnInLinkTable").ToString)
    '            Loop
    '        End Using
    '    Catch ex As Exception
    '        MsgBox("Не удалось получить вычисляемые столбцы. Ошибка:" + ex.Message)
    '    End Try

    'End Sub
    'Private Function IsColumnFireFillLinks(ByVal ColIndex As Integer, ByVal NomTable As Integer) As Boolean
    '    Dim s As String
    '    Dim dgw As DataGridView
    '    Dim dRows As DataRow()

    '    dgw = GetdgwByNom(NomTable)
    '    s = dgw.Columns(ColIndex).Name
    '    dRows = dtColAttr.Select("Table='" + mTable(NomTable) + "' AND ColumnName='" + s + "' And IsFireFillLinkColumns=True")
    '    Return (dRows.Length > 0)

    'End Function


End Module
