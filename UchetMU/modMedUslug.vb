Imports System.Data.OleDb
Imports System.Windows.Forms

Module modMedUslug
    Public AppName As String = "MedService"

    Public Sub AddStatCalc(ByVal ID_Calc As Integer, ByRef frm As frmList)
        Dim dRows As DataRow()
        Dim adap As New OleDbDataAdapter
        dRows = frm.ds2.Select("ID_Calc=" + ID_Calc.ToString)
        If dRows.Length > 0 Then Exit Sub
        adap.InsertCommand = New OleDbCommand("INSERT INTO " + frm.TableToUpdate(2) + " (ID_Calc,ID_StatCalc) SELECT " + ID_Calc.ToString + ", ID As ID_StatCalc FROM D_StatCalc", Connection)
        adap.InsertCommand.ExecuteNonQuery()
        'adapter2.FillLoadOption = LoadOption.OverwriteChanges
        frm.IsAddRowByUser = False
        frm.ds2.Clear()
        frm.adapter2.Fill(frm.ds2)
        frm.IsAddRowByUser = True
        'bindingSource2.ResetBindings(False)
    End Sub

    Public Function GetNomTableForLinkForm(ByVal TabName As String) As Integer
        If TabName = "Q_SF" Or TabName Like "Q_SF_Beznal*" Or TabName = "Q_SF_DMS" Or TabName = "Q_SF_Stom" Then
            Return 1
        Else
            Return 2
        End If
    End Function

    Public Function GetIDColumn(ByVal TableName As String, ByVal ColName As String) As String
        If ColName.ToLower = "ограничение_по_сумме" Or ColName.ToLower = "ограничение_по_процентам" Then
            Return "ФИО_пациента"
        Else
            Return "ID"
        End If
    End Function

    Public Sub AddingRowBegin(ByVal NomTable As Integer, ByVal newID As Integer, ByRef frm As frmList)
        Dim adap As New OleDbDataAdapter
        Dim sql As String
        Dim i As Integer

        If frm.mTable(NomTable) = "Q_Calc" Then AddStatCalc(newID, frm)
        If frm.mTable(NomTable) = "Q_MedUslug_NoCalc" Then
            If frm.dgwList.SelectedCells.Count = 0 Then Exit Sub
            i = frm.dgwList.Item("ID", frm.dgwList.SelectedCells(0).RowIndex).Value
            adap.InsertCommand = New OleDbCommand("INSERT INTO D_MedUslug_Podr (ID_Uslugi,ID_Podr) VALUES (" + newID.ToString + "," + i.ToString + ")", Connection)
            adap.InsertCommand.ExecuteNonQuery()
            'i = frm.dgwList.Item("ID", frm.dgwList.SelectedCells(0).RowIndex).Value
            sql = "INSERT INTO D_KodMedUslug (ID_Uslugi,ID_Year) VALUES (" + newID.ToString + "," + frmMDI.iYear.ToString + ")"
            Debug.Print(sql)
            adap.InsertCommand = New OleDbCommand(sql, Connection)
            adap.InsertCommand.ExecuteNonQuery()
            i = frm.dgwList.Item("Код_подразделения", frm.dgwList.SelectedCells(0).RowIndex).Value
            If i = 1 Or i = 20 Then
                i = 21 - i
                Dim command As New OleDbCommand("SELECT * FROM D_Podr WHERE Код_подразделения=" + i.ToString, Connection)
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read Then
                    i = reader("ID")
                    adap.InsertCommand = New OleDbCommand("INSERT INTO D_MedUslug_Podr (ID_Uslugi,ID_Podr) VALUES (" + newID.ToString + "," + i.ToString + ")", Connection)
                    'On Error Resume Next
                    adap.InsertCommand.ExecuteNonQuery()
                End If

            End If
        End If
        i = frm.dgwList.Rows.Count - 1
        If i < 0 Then Exit Sub
        If frm.mTable(NomTable) = "Q_Uslugi" Then
            i = frm.dgwList.Rows.Count - 1
            If i < 0 Then Exit Sub
            Dim ColName As String
            sql = "UPDATE R_Uslugi SET Дата_оказания_услуги=?, ID_VidStrah=? WHERE ID=" + newID.ToString
            adap.UpdateCommand = New OleDbCommand(sql, Connection)
            ColName = "Дата_оказания_услуги"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value
            ColName = "ID_VidStrah"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value
            adap.UpdateCommand.ExecuteNonQuery()
        ElseIf frm.mTable(NomTable) = "Q_DMS" Or frm.mTable(NomTable) = "Q_Stom" Then
            Dim ColName As String
            sql = "UPDATE " + frm.TableToUpdate(NomTable) + " SET Дата_посещения=?, ФИО_пациента=?, Номер_полиса=?,Дата_рождения=?,Диагноз_МКБ=? WHERE ID=" + newID.ToString
            adap.UpdateCommand = New OleDbCommand(sql, Connection)
            ColName = "Дата_посещения"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "ФИО_пациента"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "Номер_полиса"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "Дата_рождения"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value
            ColName = "Диагноз_МКБ"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value
            adap.UpdateCommand.ExecuteNonQuery()
        ElseIf frm.mTable(NomTable) = "Q_SoldSt" Then
            Dim ColName As String
            sql = "UPDATE " + frm.TableToUpdate(NomTable) + " SET Дата_начала_лечения=?, Дата_окончания_лечения=?, ФИО_пациента=?, Диагноз_МКБ=? WHERE ID=" + newID.ToString
            adap.UpdateCommand = New OleDbCommand(sql, Connection)
            ColName = "Дата_начала_лечения"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "Дата_окончания_лечения"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "ФИО_пациента"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value

            ColName = "Диагноз_МКБ"
            adap.UpdateCommand.Parameters.Add(
                        ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
            adap.UpdateCommand.Parameters(ColName).Value = frm.dgwList.Item(ColName, i).Value
            adap.UpdateCommand.ExecuteNonQuery()
        End If

    End Sub

    Public Function GetYearByDate(ByVal d As Date) As Integer
        Dim dGran As New Date(2009, 2, 18)
        If d <= dGran Then
            Return 2008
        Else
            Return 2009
        End If
    End Function

    Public Sub AddingRowEnd(ByVal NomTable As Integer, ByVal newID As Integer, ByRef frm As frmList)
        Dim adap As New OleDbDataAdapter
        Dim i, j As Integer

        If frm.mTable(NomTable) = "Q_Uslugi" Then
            i = frm.dgwList.Rows.Count - 1
            If i < 1 Then Exit Sub
            j = frm.dgwList.Columns("ID_VidStrah").Index
            frm.RefreshComboBox(1, j + 1, i)
        ElseIf frm.mTable(NomTable) = "Q_DMS" Or frm.mTable(NomTable) = "Q_Stom" Then
            i = frm.dgwList.Rows.Count - 1
            If i < 1 Then Exit Sub
        End If

    End Sub

    Public Sub PrepareLinkForm(ByRef LinkForm As frmList, ByRef frm As frmList)
        'If LinkForm.TableName = "Q_Customers" Then
        '    frm.dgwList.Columns("D_Customers").Visible = False
        '    i = frm.dgwList.Columns("D_Customers").Index
        '    frm.dgwSumma.Columns(i).Visible = False
        'End If

    End Sub

    Public Sub UpdateDependCell(ByVal Table As String, ByVal Col As Integer, ByVal Row As Integer, ByRef dgw As DataGridView)
        'Dim Cell As DataGridViewComboBoxCell
        'If Table = "Q_Uslugi" And dgw.Columns(Col).Name = "ID_Customer" Then
        '    'If IsDBNull(dgw.Item("Плательщик", Row).Value) Then
        '    '    Cell = dgw.Item("D_Customers", Row)
        '    '    If (Cell Is Nothing) Then
        '    '        dgw.Item("Плательщик", Row).Value = GetStrValue("select Краткое_наименование from D_Customers where ID=" + dgw.Item("ID", Row).Value)
        '    '        Exit Sub
        '    '    ElseIf Cell.Value = "" Then
        '    '        dgw.Item("Плательщик", Row).Value = GetStrValue("select Краткое_наименование from D_Customers where ID=" + dgw.Item("ID_Customer", Row).Value.ToString)
        '    '    Else
        '    '        dgw.Item("Плательщик", Row).Value = Cell.FormattedValue
        '    '    End If
        '    'End If
        'End If
    End Sub

    Public Function GetStrValue(ByVal sql As String) As String
        Dim Rs As OleDbDataReader
        On Error GoTo Err_h

        Dim command As New OleDbCommand(sql, Connection)
        Rs = command.ExecuteReader()
        If Not Rs.Read Then
            Return ""
        Else
            Return CStr(Rs(0).Value)
        End If
        Exit Function
Err_h:
        Return ""
    End Function

    Public Sub UpdatePodr(ByVal TableName As String)
        On Error GoTo Err_h
        Dim ColName, sql As String
        Dim adap As New OleDbDataAdapter
        Dim i As Integer

        sql = "UPDATE " + TableName + " SET " + TableName + ".ID_Podr=? WHERE ID_Uslugi IN (SELECT ID_Uslugi FROM D_MedUslug_Podr WHERE ID_Podr=?)" ' AND D_MedUslug_Podr.ID_Podr<>33"
        adap.UpdateCommand = New OleDbCommand(sql, Connection)
        'ColName = "Дата_оказания_услуги"
        adap.UpdateCommand.Parameters.Add( _
                    "ID_Podr", OleDbType.BigInt, 10)
        adap.UpdateCommand.Parameters.Add( _
                    "ID_Podr_2", OleDbType.BigInt, 10)
        For i = 18 To 41
            If i <> 33 And i <> 32 Then
                adap.UpdateCommand.Parameters("ID_Podr").Value = i
                adap.UpdateCommand.Parameters("ID_Podr_2").Value = i
                adap.UpdateCommand.ExecuteNonQuery()
            End If
        Next i
        Exit Sub
Err_h:
        MsgBox(Err.Description)
    End Sub
End Module
