'Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Data
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports Word = Microsoft.Office.Interop.Word

Module modReport
    Public objWord As Word.Application
    Public Doc As Word.Document
    Public objExcel As Excel.Application
    Public xlBook As Excel.Workbook
    Public xlSheet As Excel.Worksheet
    Public rGroup1, rGroup2 As Excel.Range
    Public rRow As Excel.Range
    'Public begin As Point
    Public PathRep, Query, Templ As String
    Public IsDebug As Boolean = False
    Private NomPP As Integer = 1
    Private reader As DataTableReader
    Private dat As DataTable
    Private GrFieldCount As Integer
    Private ID_PerFrom, ID_PerTo, ID_VidReport As Integer
    Private SenderName As String
    Public Filter As String
    Public Header As String
    Public dBegin As Date
    Public dEnd As Date
    Private IsNomPP As Boolean
    Private sFilterPeriod, sFIO As String
    Private sPeriod As String = ""
    Private sTitle As String = ""
    Private ID_Podr, ID_SF As Integer

    Public Function GetExcel() As Boolean
        On Error GoTo Excel_OLE_err
        Dim IsExcelOpen As Boolean

        objExcel = GetObject(, "Excel.Application")
        'If objExcel.WindowState <> wdWindowStateMinimize Then objExcel.WindowState = wdWindowStateMinimize
        IsExcelOpen = True
        GetExcel = True
        Exit Function

Excel_OLE_err:
        Select Case Err.Number
            Case 429
                IsExcelOpen = False
                Resume Excel_OLE_continue
            Case Else
                MsgBox("Невозможно запустить MS Excel на Вашем компьютере." & vbCrLf & "Ошибка " & Err.Number & ": " & Err.Description & ".", vbExclamation, "Ошибка")
                GetExcel = False
        End Select

Excel_OLE_continue:
        If IsExcelOpen = False Then
            objExcel = CreateObject("Excel.Application")
            GetExcel = True
        End If
    End Function

    Public Function GetWord() As Boolean
        On Error GoTo Word_OLE_err
        Dim IsWordOpen As Boolean

        objWord = GetObject(, "Word.Application")
        'If objExcel.WindowState <> wdWindowStateMinimize Then objExcel.WindowState = wdWindowStateMinimize
        IsWordOpen = True
        GetWord = True
        Exit Function

Word_OLE_err:
        Select Case Err.Number
            Case 429
                IsWordOpen = False
                Resume Word_OLE_continue
            Case Else
                MsgBox("Невозможно запустить MS Word на Вашем компьютере." & vbCrLf & "Ошибка " & Err.Number & ": " & Err.Description & ".", vbExclamation, "Ошибка")
                GetWord = False
        End Select

Word_OLE_continue:
        If IsWordOpen = False Then
            objWord = CreateObject("Word.Application")
            GetWord = True
        End If
    End Function

    'Public Sub SimpleMessage()
    '    Crypt.EXECryptor_CRYPT_START
    '    MsgBox "Работает просто."
    '    Crypt.EXECryptor_CRYPT_END
    '
    'End Sub
    '
    'Public Sub Message()
    '        Crypt.EXECryptor_CRYPT_START
    '        MsgBox "Работает! Осталось " & CStr(Crypt.GetTrialRunsLeft(100))
    '        If IsTrial Then
    '            MsgBox "Работает! Осталось " & CStr(Crypt.GetTrialDaysLeft(-30))
    '        Else
    '            MsgBox "Работает. Период истек."
    '        End If
    '        Crypt.EXECryptor_CRYPT_END

    'End Sub

    Private Function GetExcelSheet() As Boolean
        On Error GoTo Err_h
        PathRep = CurDir() & "\Resources"
        If Not GetExcel() Then
            MsgBox("Не удалось открыть MS Excel.")
            Exit Function
        End If
        PathRep += "\" + Templ
        xlBook = objExcel.Workbooks.Add(PathRep) 'ошибка здесь

        xlSheet = xlBook.Sheets(1)
        If IsDebug Then objExcel.Visible = True
        Return True
Err_h:
        ErrMess(Err.Description, "GetExcelSheet")
        'MsgBox("Не найден файл шаблона " + Templ)
        Return False
    End Function

    Private Function GetWordDoc() As Boolean
        On Error GoTo Err_h

        PathRep = CurDir() & "\Resources"
        If Not GetWord() Then
            MsgBox("Не удалось открыть MS Word.")
            Exit Function
        End If
        PathRep += "\" + Templ
        Doc = objWord.Documents.Add(PathRep)

        If IsDebug Then objWord.Visible = True
        Return True
Err_h:
        Return False
    End Function

    Public Sub CreateDW(ByVal sender As System.Object)
        If Connection.State = ConnectionState.Closed Then Connection.Open()
        Dim ComDel As New OleDbCommand("delete from R_SumByStatNal", Connection)
        ComDel.ExecuteNonQuery()
        Dim ComIns As New OleDbCommand("execute Ins_R_SumByStatNal", Connection)
        ComIns.ExecuteNonQuery()
    End Sub

    Public Sub CreateReport(ByVal sender As System.Object)
        Dim RC As Boolean
        Dim dt As DataTable


        On Error GoTo Err_h

        If Filter = "" Then Filter = "1=1"
        If Connection.State = ConnectionState.Closed Then Connection.Open()
        SenderName = sender.Name
        dt = Populate("select * from D_Reports where Menu='" + sender.Name + "'")
        reader = New System.Data.DataTableReader(dt)

        Do While reader.Read
            Query = reader("Query").ToString
            Templ = reader("Templ").ToString
            Exit Do
        Loop
        If (GetParamForReport(Query) = False) Then Exit Sub
        If Query = "Q_ReportBeznal" Or Query = "Q_IncomesByStat" Or Query = "Q_Zakaz" Then
            RC = CreateReportByQuery(Query, SenderName, Nothing, Nothing)
            If RC Then SaveReport()
        Else  'Для других отчётов запускаем второй поток
            Dim frm As New frmReport
            frm.ReportQuery = Query
            frm.ReportTitle = Query
            frm.SenderName = SenderName
            frm.Show()
        End If

        'Select Case Query
        '    Case "Q_Preiskurant"
        '        If SenderName = "mnuPreiskurantBas" Then
        '            '    RC = Rep_Preiskurant("Код_подразделения<>20 AND Код_услуги_по_номенклатуре<>''")
        '            RC = Rep_Preiskurant("Код_услуги_по_номенклатуре<>''")
        '        ElseIf SenderName = "mnuPreiskurantOther" Then
        '            RC = Rep_Preiskurant("IsNull(Код_услуги_по_номенклатуре)=True")
        '        End If
        '    Case "Q_ReestrUslug"
        '        RC = Rep_ReestrUslug()
        '    Case "Q_IncomesByPodrAndStat"
        '        RC = Rep_IncomesByPodrAndStat()
        '    Case "Q_Akt"
        '        RC = Rep_Akt(SenderName)
        '        If RC Then objWord.Visible = True
        '        objWord = Nothing
        '        Exit Sub
        '    Case "Q_ReportBeznal"
        '        RC = Rep_ReestrSF()
        '    Case "Q_Zakaz"
        '        RC = Rep_Zakaz()
        '    Case "Q_IncomesByStat"
        '        RC = Rep_IncomesByStat()
        '    Case "Q_UslugiByStat"
        '        RC = Rep_ReestrUslugByStat()
        'End Select
        'If RC And objExcel.Visible = False Then
        '    If frmSave.ShowDialog(frmMDI) = System.Windows.Forms.DialogResult.OK Then
        '        If frmSave.optOpen.Checked Then
        '            If Not objExcel.Visible Then
        '                objExcel.Visible = True
        '            End If
        '            objExcel.WindowState = Excel.XlWindowState.xlMaximized
        '        Else
        '            With dlgSave
        '                '.InitialDirectory = "c:\"
        '                .Filter = "Файлы MS Excel (*.xls)|*.xls|Все файлы (*.*)|*.*"
        '                .FilterIndex = 1
        '                .RestoreDirectory = True
        '            End With
        '            If dlgSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '                xlBook.SaveAs(dlgSave.FileName)
        '            End If
        '        End If
        '    End If
        'End If
        'objExcel = Nothing
        'frm.Visible = False
        Exit Sub
Err_h:
        ErrMess(Err.Description, "CreateReport")
    End Sub

    Public Function Rep_Preiskurant(ByVal Filter As String, ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim Podr As String = ""
        Dim s As String
        Dim row As Excel.Range
        Dim rowTempl, GroupTempl As Integer
        Dim b As Boolean
        'On Error GoTo Err_h
        On Error Resume Next

        'If frmFilterDataSimple.ShowDialog() = DialogResult.Cancel Then
        '    Return False
        'End If
        If Not GetExcelSheet() Then
            Return False
        End If
        dEnd = dBegin
        dat = PopulateInterval("select * from " + Query + " where Дата_начала_действия<=? AND Дата_окончания_действия>=? AND " + Filter, dBegin, dBegin)
        Header = " на дату " + FormatDateTime(dBegin.Date, DateFormat.ShortDate)
        s = xlSheet.Range("HReport").Cells(1, 1).Value()
        s += Header
        xlSheet.Range("HReport").Cells(1, 1).Value = s

        IsNomPP = False
        GroupTempl = xlSheet.Range("HGRoup1").Row
        'rGroup2 = xlSheet.Range("HGRoup2")
        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        reader = New System.Data.DataTableReader(dat)
        Do While reader.Read
            If Podr = "" Then
                xlSheet.Rows(GroupTempl).Cells(1, 1) = reader("Код_подразделения").ToString + ". " + reader("Подразделение").ToString
            ElseIf Podr <> reader("Подразделение") Then
                GroupTempl = xlSheet.Range("HGRoup1").Row
                xlSheet.Rows(GroupTempl).Copy()
                InsertRow(row)
                row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
                row.ClearContents()
                row.Cells(1, 1) = reader("Код_подразделения").ToString + ". " + reader("Подразделение").ToString
                ShiftRow(row, 1)
                b = row.Font.Bold
                Podr = reader("Подразделение")
                'NomPodr += 1
            End If
            b = row.Font.Bold
            xlSheet.Rows(rowTempl).Copy()
            InsertRow(row)
            row.Font.Bold = b
            FillRow(row, "Код_подразделения,Подразделение,Дата_начала_действия,Дата_окончания_действия")
            ShiftRow(row, 1)
            Podr = reader("Подразделение")
            If bw.CancellationPending Then
                e.Cancel = True
                CancelReport()
                Return False
            End If
        Loop
        'row.Delete()
        Return True
        Exit Function
Err_h:
        ErrMess(Err.Description, "Rep_Preiskurant")
        Return False
    End Function

    Public Function Rep_ReestrUslugByStat(ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim Podr As String = ""
        'Dim Per As Period
        Dim s, QuerySum, QuerySumAll, Query2, Query, sqlSumAll, DelSQL As String
        Dim QueryExecute() As String
        Dim row As Excel.Range
        Dim rowTempl, GroupTempl, GroupEndTempl As Integer
        Dim bool As Boolean = True
        Dim MaxStatCalcID As Integer
        'Dim frm As New frmRepReestrByStat
        Dim ID_Uslugi, col, i, rowindex As Integer
        On Error GoTo Err_h

        ReDim QueryExecute(0 To 5)
        If sFilterPeriod = "" Then Return False
        Query = "R_SumByStat"
        DelSQL = "DELETE FROM R_SumByStat WHERE Vid=" + ID_VidReport.ToString + " AND ID_Period IN (" + sFilterPeriod + ")"
        Select Case ID_VidReport
            Case -1
                MsgBox("Необходимо выбрать вид отчета.")
                Exit Function
            Case 0 'наличный расчет
                QueryExecute(ID_VidReport) = "EXECUTE Ins_R_SumByStatNal"
                'Query = "Q_SumByStatNal"
                QuerySum = "Q_IncomesByPodrAndStatNal"
                QuerySumAll = "Q_IncomesByStatNal"
            Case 1 'безналичный расчет
                QueryExecute(ID_VidReport) = "EXECUTE Ins_R_SumByStatBezNal"
                'Query = "Q_SumByStatBezNal"
                QuerySum = "Q_IncomesByPodrAndStatBeznal"
                QuerySumAll = "Q_IncomesByStatBeznal"
            Case 2 'военнослужащие
                QueryExecute(ID_VidReport) = "EXECUTE Ins_R_SumByStatSoldSt"
                'Query = "Q_SumByStatSoldSt"
                QuerySum = "Q_IncomesByPodrAndStatSold"
                QuerySumAll = "Q_IncomesByStatSold"
            Case 3 'ДМС
                QueryExecute(ID_VidReport) = "EXECUTE Ins_R_SumByStatDMS"
                'Query = "Q_SumByStatDMS"
                QuerySum = "Q_IncomesByPodrAndStatDMS"
                QuerySumAll = "Q_IncomesByStatDMS"
            Case 4 'УМО
                QueryExecute(ID_VidReport) = "EXECUTE Ins_R_SumByStatUMO"
                'Query = "Q_SumByStatUMO"
                QuerySum = "Q_IncomesByPodrAndStatUMO"
                QuerySumAll = "Q_IncomesByStatUMO"
            Case 5 'Все
                DelSQL = "DELETE FROM R_SumByStat WHERE ID_Period IN (" + sFilterPeriod + ")"
                QueryExecute(0) = "EXECUTE Ins_R_SumByStatNal"
                QueryExecute(1) = "EXECUTE Ins_R_SumByStatBezNal"
                QueryExecute(2) = "EXECUTE Ins_R_SumByStatSoldSt"
                QueryExecute(3) = "EXECUTE Ins_R_SumByStatDMS"
                QueryExecute(4) = "EXECUTE Ins_R_SumByStatUMO"
                Query = "Q_SumByStatAll"
                QuerySum = "Q_IncomesByPodrAndStat"
                QuerySumAll = "Q_IncomesByStat"
        End Select
        If Connection.State = ConnectionState.Closed Then Connection.Open()

        Dim ComDel As New OleDbCommand(DelSQL, Connection)
        ComDel.ExecuteNonQuery()
        Dim comPer As New OleDbCommand("SELECT ID FROM D_Periods WHERE ID IN (" + sFilterPeriod + ")", Connection)
        Dim readerPer As OleDbDataReader = comPer.ExecuteReader()
        If ID_VidReport = 5 Then
            Do While readerPer.Read
                For i = 0 To 4
                    Dim ComIns As New OleDbCommand(QueryExecute(i) + " " + readerPer("ID").ToString, Connection)
                    ComIns.ExecuteNonQuery()
                    ComIns.Dispose()
                Next i
            Loop
        Else
            Do While readerPer.Read
                Dim ComIns As New OleDbCommand(QueryExecute(ID_VidReport) + " " + readerPer("ID").ToString, Connection)
                ComIns.ExecuteNonQuery()
                ComIns.Dispose()
            Loop
        End If
        If Not GetExcelSheet() Then
            Return False
        End If
        'для отладки
        'objExcel.Visible = True
        xlSheet.Range("HReport").Cells(1, 1).Value = sTitle
        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        IsNomPP = False
        'Dim ID_Podr As Integer
        GroupTempl = xlSheet.Range("HGRoup1").Row
        Dim sqlSum As String
        If ID_Podr = -1 Then
            Query2 = "select ID_Uslugi, ID_Calc, ID_StatCalc, ID_Podr, ID_KodUslugi, Код_услуги, Код_услуги_по_номенклатуре, Услуга, Единица_измерения, SUM(Количество) AS Количество, Цена, SUM(Сумма) AS Сумма FROM " + Query + " WHERE Vid=" + ID_VidReport.ToString + " AND ID_Period IN (" + sFilterPeriod + ") AND ID_Podr=" + ID_Podr.ToString + " GROUP BY ID_Uslugi, ID_Calc, ID_StatCalc, ID_Podr, ID_KodUslugi, Код_услуги, Код_услуги_по_номенклатуре, Услуга, Единица_измерения,Цена ORDER BY Код_услуги, ID_Calc, ID_StatCalc"
            sqlSum = "SELECT ID_Podr, Код_подразделения, Подразделение, ID_StatCalc, SUM(Сумма) AS Сумма FROM " + QuerySum + " WHERE ID_Period IN (" + sFilterPeriod + ") AND ID_Podr=" + ID_Podr.ToString + " GROUP BY ID_Podr, Код_подразделения, Подразделение, ID_StatCalc ORDER BY  Код_подразделения, ID_StatCalc"
            sqlSumAll = sqlSum
            'sqlSum = "SELECT * FROM " + QuerySum + " WHERE ID_Period=" + frm.Per.ID.ToString + " AND ID_Podr=" + ID_Podr.ToString + " ORDER BY Код_подразделения, ID_StatCalc"
            'sqlSumAll = "SELECT * FROM " + QuerySum + " WHERE ID_Period=" + frm.Per.ID.ToString + " AND ID_Podr=" + ID_Podr.ToString
        Else
            Query2 = "select ID_Uslugi, ID_Calc, ID_StatCalc, ID_Podr, ID_KodUslugi, Код_услуги, Код_услуги_по_номенклатуре, Услуга, Единица_измерения, SUM(Количество) AS Количество, Цена, SUM(Сумма) AS Сумма FROM " + Query + " WHERE Vid=" + ID_VidReport.ToString + " AND ID_Podr=? AND ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_Uslugi, ID_Calc, ID_StatCalc, ID_Podr, ID_KodUslugi, Код_услуги, Код_услуги_по_номенклатуре, Услуга, Единица_измерения,Цена  ORDER BY Код_услуги, ID_Calc, ID_StatCalc"
            sqlSum = "SELECT ID_Podr, Код_подразделения, Подразделение, ID_StatCalc, SUM(Сумма) AS Сумма FROM " + QuerySum + " WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_Podr, Код_подразделения, Подразделение, ID_StatCalc ORDER BY Код_подразделения, ID_StatCalc"
            sqlSumAll = "SELECT  ID_StatCalc, SUM(Сумма) AS Сумма FROM " + QuerySumAll + " WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_StatCalc ORDER BY ID_StatCalc"
            'sqlSumAll = "SELECT * FROM " + QuerySumAll + " WHERE ID_Period=" + frm.Per.ID.ToString
            ID_Podr = 0
        End If
        Debug.Print(sqlSum)
        Debug.Print(sqlSumAll)
        Dim command2 As New OleDbCommand(sqlSum, Connection)
        Dim readerSum As OleDbDataReader = command2.ExecuteReader()
        Dim Query2_ As String
        If readerSum.Read() Then
            bool = True
            Do While bool
                ID_Podr = readerSum("ID_Podr")
                If xlSheet.Rows(GroupTempl).Cells(1, 1).Value Is Nothing Then
                    xlSheet.Rows(GroupTempl).Cells(1, 1).Value = IIf(readerSum("Код_подразделения") = 100, "", readerSum("Код_подразделения").ToString + ".") + readerSum("Подразделение").ToString
                Else
                    xlSheet.Rows(GroupTempl).Copy()
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
                    row.ClearContents()
                    row.Cells(1, 1).Value = IIf(readerSum("Код_подразделения") = 100, "", readerSum("Код_подразделения").ToString + ".") + readerSum("Подразделение").ToString
                    'NomPodr += 1
                    ShiftRow(row, 2)
                End If
                Query2_ = Replace(Query2, "?", ID_Podr.ToString)
                Debug.Print(Query2_)
                Dim command As New OleDbCommand(Query2_, Connection)
                Dim reader As OleDbDataReader = command.ExecuteReader()
                Dim ID_Calc As Integer
                If reader.Read() Then
                    bool = True
                    Do While (bool)
                        If Not IsDBNull(reader("Код_услуги")) Then Debug.Print(reader("Код_услуги"))
                        If xlSheet.Rows(rowTempl).Cells(1, 1).Value Is Nothing Then
                            row = xlSheet.Rows(rowTempl)
                        Else
                            xlSheet.Rows(rowTempl).Copy()
                            InsertRow(row)
                            ShiftRow(row, 2)
                            ClearAllCells(row)
                            row.Font.Bold = 0
                        End If
                        FillRowOleDB(row, reader, "Код_подразделения,Подразделение,Сумма")
                        ID_Calc = reader("ID_Calc")
                        col = 8
                        Do While (ID_Calc = reader("ID_Calc"))
                            If col = 14 Then col += 1
                            row.Cells(1, col).Value = reader("Сумма").ToString
                            Debug.Print(reader("ID_StatCalc").ToString)
                            col += 1
                            If Not reader.Read() Then
                                bool = False
                                Exit Do
                            End If
                        Loop
                        ShiftRow(row, 2)
                    Loop
                Else
                    'ShiftRow(row, -2)
                    i = row.Row
                    row.Delete()
                    If i < xlSheet.Range("FGRoup1").Row Then
                        i = xlSheet.Range("FGRoup1").Row - 1
                        xlSheet.Range("FGRoup1").Name.Name = ""
                        Dim ran As Excel.Range
                        ran = xlSheet.Rows(i)
                        ran.Name = "FGRoup1"
                        i = xlSheet.Range("FReport").Row - 1
                        xlSheet.Range("FReport").Name.Name = ""
                        ran = xlSheet.Rows(i)
                        ran.Name = "FReport"
                        If ID_VidReport = -1 Then
                            row = xlSheet.Rows(xlSheet.Range("FGRoup1").Row)
                        Else
                            row = xlSheet.Rows(xlSheet.Range("FGRoup1").Row + 1)
                        End If
                    Else
                        row = xlSheet.Rows(i)
                    End If
                End If
                reader.Close()
                reader = Nothing
                command.Dispose()
                command = Nothing
                GroupEndTempl = xlSheet.Range("FGRoup1").Row
                If row.Row > GroupEndTempl Then
                    xlSheet.Rows(GroupEndTempl).Copy()
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
                    ClearAllCells(row)
                    row.Cells(1, 2).Value = "Итого по"
                Else
                    row = xlSheet.Rows(GroupEndTempl)
                End If
                row.Cells(1, 3).Value = IIf(readerSum("Код_подразделения") = 100, "", readerSum("Код_подразделения").ToString + ".") + readerSum("Подразделение").ToString
                ID_Podr = readerSum("ID_Podr")
                col = 8
                bool = True
                Do While (ID_Podr = readerSum("ID_Podr"))
                    If col = 14 Then col += 1
                    If Not IsDBNull(readerSum("Сумма")) Then row.Cells(1, col).Value = readerSum("Сумма").ToString
                    col += 1
                    If Not readerSum.Read() Then
                        bool = False
                        Exit Do
                    End If
                Loop
                ShiftRow(row, 2)
                If bw.CancellationPending Then
                    e.Cancel = True
                    CancelReport()
                    Return False
                End If
            Loop
        End If
        readerSum.Close()
        readerSum = Nothing
        command2.Dispose()
        command2 = Nothing
        row = xlSheet.Range("FReport")
        Dim command3 As New OleDbCommand(sqlSumAll, Connection)
        Dim readerSumAll As OleDbDataReader = command3.ExecuteReader()
        col = 8
        Do While readerSumAll.Read()
            If col = 14 Then col += 1
            If Not IsDBNull(readerSumAll("Сумма")) Then row.Cells(1, col).Value = readerSumAll("Сумма").ToString
            Debug.Print(readerSumAll("ID_StatCalc").ToString)
            col += 1
        Loop
        command3.Dispose()
        Dim comDelAll As New OleDbCommand("delete from R_SumByStat", Connection)
        comDelAll.ExecuteNonQuery()
        Return True
        Exit Function
Err_h:
        objExcel.Visible = True
        If Err.Number = 5 Then
            MsgBox("Ошибка в запросе БД.")
        Else
            ErrMess(Err.Description, "Rep_ReestrUslugByStat")
        End If
        Return False
    End Function

    Public Sub ClearAllCells(ByRef row As Excel.Range)
        Dim i As Integer
        For i = 1 To xlSheet.UsedRange.Columns.Count
            If Not row.Cells(1, i).HasFormula Then row.Cells(1, i).Value = ""
        Next i
    End Sub


    Public Function Rep_Zakaz() As Boolean
        Dim s As String
        Dim row As Excel.Range
        Dim rowTempl, SelRow As Integer
        'Dim frm As New frmFilterZakaz
        'Dim frm2 As frmList
        Dim Summa As Decimal
        'On Error GoTo Err_h
        On Error Resume Next

        'frm2 = frmMDI.ActiveMdiChild
        'If frm2.dgwList.SelectedCells.Count = 0 Then
        '    MsgBox("Выберите счет-фактуру.")
        '    Return False
        'End If
        'SelRow = frm2.dgwList.SelectedCells(0).RowIndex
        'frm.ID_Sf = frm2.dgwList.Item("ID", SelRow).Value
        'If frm.ShowDialog() = DialogResult.Cancel Then
        '    Return False
        'End If
        If Not GetExcelSheet() Then
            Return False
        End If
        Dim command As New OleDbCommand("SELECT * FROM " + Query + " WHERE ID_SF=" + ID_SF.ToString + " AND ФИО_пациента='" + sFIO + "'", Connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        If Not reader.Read Then
            Return False
        End If
        s = reader("Страховая_компания").ToString
        xlSheet.Range("StrahComp").Cells(1, 1).Value = s
        s = reader("Номер_полиса").ToString
        xlSheet.Range("Polis").Cells(1, 1).Value = s
        s = reader("ФИО_пациента").ToString
        xlSheet.Range("Fam").Cells(1, 1).Value = s
        s = reader("Дата_рождения").ToString
        xlSheet.Range("DataRojd").Cells(1, 1).Value = s

        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        Do
            xlSheet.Rows(rowTempl).Copy()
            InsertRow(row)
            ShiftRow(row, 2)
            row.Font.Bold = 0
            s = reader("Диагноз_МКБ").ToString
            row.Cells(1, 1).Value = s
            s = reader("Код_услуги").ToString
            row.Cells(1, 4).Value = s
            s = reader("Услуга").ToString
            row.Cells(1, 7).Value = s
            s = reader("Цена").ToString
            row.Cells(1, 18).Value = s
            s = reader("Количество_оказанной_услуги").ToString
            row.Cells(1, 20).Value = s
            s = reader("Сумма").ToString
            row.Cells(1, 21).Value = s
        Loop Until Not reader.Read
        ShiftRow(row, 3)
        row.Delete()
        Dim command2 As New OleDbCommand("SELECT SUM(Сумма) AS Сумма FROM " + Query + " WHERE ID_SF=" + ID_SF.ToString + " AND ФИО_пациента='" + sFIO + "'", Connection)
        Dim reader2 As OleDbDataReader = command2.ExecuteReader()
        If Not reader2.Read Then
            Return False
        End If
        Summa = reader2("Сумма")
        xlSheet.Range("Itogo").Cells(1, 1).Value = Summa.ToString
        xlSheet.Range("SummaProp").Cells(1, 1).Value = GetSummaMoneyString(Summa)

        Return True
        Exit Function
Err_h:
        ErrMess(Err.Description, "Rep_Preiskurant")
        Return False
    End Function

    Public Function Rep_Akt(ByVal SenderName As String) As Boolean
        Dim frm As frmList
        Dim SelRow, ID_SF, ID_StrahComp As Integer
        Dim s As String
        Dim d As Date
        'On Error GoTo Err_h 
        On Error Resume Next

        frm = frmMDI.ActiveMdiChild
        If frm.dgwList.SelectedCells.Count = 0 Then
            MsgBox("Выберите счет-фактуру.")
        End If
        SelRow = frm.dgwList.SelectedCells(0).RowIndex
        ID_SF = frm.dgwList.Item("ID", SelRow).Value
        If frm.ID_VidSF = 3 Then
            Query = "Q_Akt_DMS"
        ElseIf frm.ID_VidSF = 6 Then
            Query = "Q_Akt_Stom"
        End If
        Dim command3 As New OleDbCommand("SELECT * FROM " + Query + " WHERE ID_SF=" + ID_SF.ToString, Connection)
        Dim reader3 As OleDbDataReader = command3.ExecuteReader()
        If Not reader3.Read Then
            Exit Function
        End If
        ID_StrahComp = reader3("ID_StrahComp")
        If Not IsDBNull(reader3("Шаблон_акта")) Then Templ = reader3("Шаблон_акта")
        If Not GetWordDoc() Then
            Return False
        End If
        If Doc.Bookmarks.Exists("NomDog") Then
            s = reader3("Номер_договора").ToString
            Doc.Bookmarks.Item("NomDog").Range.Text = s
        End If
        If Not IsDBNull(reader3("Дата_договора")) Then
            d = reader3("Дата_договора").ToString
            s = DateToStr(d)
            If Doc.Bookmarks.Exists("DataDog") Then Doc.Bookmarks.Item("DataDog").Range.Text = s
            If Doc.Bookmarks.Exists("DataDog2") Then Doc.Bookmarks.Item("DataDog2").Range.Text = s
        End If
        s = reader3("Номер").ToString
        Doc.Bookmarks.Item("NomSF").Range.Text = s
        If Doc.Bookmarks.Exists("NomSF2") Then Doc.Bookmarks.Item("NomSF2").Range.Text = s
        If Not IsDBNull(reader3("Дата")) Then
            d = reader3("Дата")
            s = DateToStr(d)
            If Doc.Bookmarks.Exists("DataSF") Then Doc.Bookmarks.Item("DataSF").Range.Text = s
            If Doc.Bookmarks.Exists("DataSF2") Then Doc.Bookmarks.Item("DataSF2").Range.Text = s
        End If
        s = reader3("Страховая_компания").ToString
        If Doc.Bookmarks.Exists("StrahComp") Then Doc.Bookmarks.Item("StrahComp").Range.Text = s
        If Doc.Bookmarks.Exists("StrahComp2") Then Doc.Bookmarks.Item("StrahComp2").Range.Text = s
        If Doc.Bookmarks.Exists("StrahComp3") Then Doc.Bookmarks.Item("StrahComp3").Range.Text = s
        s = reader3("Сумма_к_оплате").ToString
        Doc.Bookmarks.Item("SummaSF").Range.Text = s
        If Templ = "Акт по Согазу.dot" Then
            Dim kop As Integer
            kop = 100 * (s - Int(s))
            s = GetSummaString(Int(s), 1, "", "", "")
            's = GetSummaString(Int(s), 1, "рубль", "рубля", "рублей")
            Doc.Bookmarks.Item("SummaProp").Range.Text = s
            Doc.Bookmarks.Item("SummaKop").Range.Text = kop.ToString
            Dim command4 As New OleDbCommand("SELECT * FROM Q_Akt_DMS_2 WHERE ID_SF=" + ID_SF.ToString, Connection)
            Dim reader4 As OleDbDataReader = command4.ExecuteReader()
            If reader4.Read Then
                s = FormatDateTime(reader4("Дата_начала"), DateFormat.ShortDate)
                Doc.Bookmarks.Item("DataNach").Range.Text = s
                s = FormatDateTime(reader4("Дата_окончания"), DateFormat.ShortDate)
                Doc.Bookmarks.Item("DataOkon").Range.Text = s
            End If
        Else
            s = GetSummaMoneyString(s)
            Doc.Bookmarks.Item("SummaSFProp").Range.Text = s
        End If

        If Templ = "Акт по Мед-Визиту.dot" Then
            Dim command4 As New OleDbCommand("SELECT * FROM Q_Akt_DMS_2 WHERE ID_SF=" + ID_SF.ToString, Connection)
            Dim reader4 As OleDbDataReader = command3.ExecuteReader()
            s = reader4("Дата_начала").ToString
            Doc.Bookmarks.Item("DataNach").Range.Text = s
            s = reader4("Дата_окончания").ToString
            Doc.Bookmarks.Item("DataOkon").Range.Text = s
            s = reader4("Количество_пациентов").ToString
            Doc.Bookmarks.Item("KolZastr").Range.Text = s
        End If
        Return True
Err_h:
        ErrMess(Err.Description, "Rep_Preiskurant")
        Return False
    End Function

    Public Function Rep_ReestrUslug(ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim s As String = ""
        Dim row As Excel.Range
        Dim rowTempl As Integer
        Dim b As Boolean
        On Error GoTo Err_h

        IsNomPP = False
        'NomPodr = 1
        'NomPP = 1

        'Dim frm As New frmFilterData
        'frm.optNoPay.Visible = False
        'frm.optAll.Location = frm.optNoPay.Location
        'If frm.ShowDialog() = DialogResult.Cancel Then
        '    Return False
        'End If
        If Not GetExcelSheet() Then
            Return False
        End If
        'objExcel.Visible = True
        dat = PopulateInterval("select * from " + Query + " where Дата_оказания_услуги>=? AND Дата_оказания_услуги<=?", dBegin, dEnd)
        Header = " c " + FormatDateTime(dBegin.Date, DateFormat.ShortDate) + " по " + FormatDateTime(dEnd.Date.ToString, DateFormat.ShortDate)
        s = xlSheet.Range("HReport").Cells(1, 1).Value()
        s += Header
        xlSheet.Range("HReport").Cells(1, 1).Value = s
        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        reader = New System.Data.DataTableReader(dat)
        Do While reader.Read
            xlSheet.Rows(rowTempl).Copy()
            InsertRow(row)
            ShiftRow(row, 2)
            row.Font.Bold = 0
            FillRow(row)
            'ShiftRow(row, 1)
            If bw.CancellationPending Then
                e.Cancel = True
                CancelReport()
                Return False
            End If
        Loop
        'row.Delete()
        Return True
        Exit Function
Err_h:
        ErrMess(Err.Description)
        Return False
    End Function

    Public Function Rep_ReestrSF() '(ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim s As String = ""
        Dim row As Excel.Range
        Dim rowTempl As Integer
        Dim b As Boolean
        Dim frm As frmList
        Dim ID_SF, ID_VidSF, SelRow, ID_StrahComp, i, j, RowNomPP As Integer
        Dim Shift As Integer = 0
        Dim ShiftSumma As Integer = 0
        Dim Customer As String = ""
        Dim sql As String = ""
        Dim ExceptColumns As String = ""
        Dim Header2 As String = ""
        Dim Summa As Decimal
        'Dim IsFirstRow As Boolean
        'Dim FIO_old As String

        On Error GoTo Err_h

        IsNomPP = True
        'NomPodr = 1
        NomPP = 1
        frm = frmMDI.ActiveMdiChild
        If frm.dgwList.SelectedCells.Count = 0 Then
            MsgBox("Выберите счет-фактуру.")
        End If
        SelRow = frm.dgwList.SelectedCells(0).RowIndex
        ID_SF = frm.dgwList.Item("ID", SelRow).Value
        Select Case frm.ID_VidSF
            Case 1 'Военнослужащие амбулаторно'
                Exit Function
            Case 2 'Военнослужащие стационарно'
                'Customer = frm.dgwList.Item("D_StrahComp", SelRow).FormattedValue
                Templ = "Реестр военнослужащие стационарно.xlt"
                dat = Populate("select * from Q_ReportSoldSt where ID_SF=" + ID_SF.ToString)
                Shift = 1
                sql = "SELECT Sum(Сумма) FROM Q_ReportSoldSt where ID_SF=" + ID_SF.ToString
                ExceptColumns = "Цена,количество_оказанной_услуги"
                ID_StrahComp = frm.dgwList.Item("ID_StrahComp", SelRow).Value
                Dim command2 As New OleDbCommand("SELECT Страховая_компания FROM D_StrahComp WHERE ID=" + ID_StrahComp.ToString, Connection)
                Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                If reader2.Read Then
                    Header = "для работающих в " + reader2(0).ToString
                End If
            Case 3 'ДМС
                Templ = "Реестр ДМС.xlt"
                Shift = 1
                sql = "SELECT Sum(Сумма_к_оплате) AS Сумма FROM Q_DMS where ID_SF=" + ID_SF.ToString
                ID_StrahComp = frm.dgwList.Item("ID_StrahComp", SelRow).Value
                Dim command2 As New OleDbCommand("SELECT Страховая_компания FROM D_StrahComp WHERE ID=" + ID_StrahComp.ToString, Connection)
                Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                If reader2.Read Then
                    Customer = reader2(0).ToString
                    Header = "лиц, застрахованных в " + reader2(0).ToString + " для оплаты медицинских услуг"
                End If
            Case 4 'безналичный расчет
                Templ = "Реестр СФ факт.xlt"
                Shift = 1
                dat = Populate("select * from " + Query + " where ID_SF=" + ID_SF.ToString)
                sql = "SELECT Sum(Сумма) FROM " + Query + " where ID_SF=" + ID_SF.ToString
                Header = "Фактическая стоимость периодического медицинского осмотра"
                Customer = frm.dgwList.Item("Заказчик", SelRow).FormattedValue
                'Dim command2 As New OleDbCommand("SELECT Страховая_компания FROM D_StrahComp WHERE ID=" + ID_StrahComp.ToString, Connection)
                'Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                'If reader2.Read Then
                Header2 = "в 2009 году. Заказчик " + Customer
                Customer = ""
                'End If
            Case 5 'безналичный расчет - проект
                Templ = "Реестр СФ.xlt"
                'Header = "РЕЕСТР к счету (счет-фактуре) № " + frm.dgwList.Item("Номер", SelRow).Value + " от " + frm.dgwList.Item("Дата", SelRow).Value
                Shift = 1
                dat = Populate("select * from " + Query + " where ID_SF=" + ID_SF.ToString)
                sql = "SELECT Sum(Сумма) FROM " + Query + " where ID_SF=" + ID_SF.ToString
                'Header = "Предварительная стоимость периодического медицинского осмотра"
                Customer = frm.dgwList.Item("Заказчик", SelRow).FormattedValue
                'Dim command2 As New OleDbCommand("SELECT Страховая_компания FROM D_StrahComp WHERE ID=" + ID_StrahComp.ToString, Connection)
                'Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                'If reader2.Read Then
                Header2 = "в 2009 году. Заказчик " + Customer
                Customer = ""
                'End If
            Case 6 'Стоматология
                Customer = frm.dgwList.Item("Страховая_компания", SelRow).FormattedValue
                Templ = "Реестр стоматология.xlt"
                Shift = 1
                sql = "SELECT Sum(Сумма_к_оплате) AS Сумма FROM Q_Stom where ID_SF=" + ID_SF.ToString
                Header = ""
        End Select
        If Not GetExcelSheet() Then
            Return False
        End If
        s += Header
        If frm.ID_VidSF < 5 Then xlSheet.Range("HReport").Cells(1, 1).Value = s
        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        If frm.ID_VidSF <> 3 And frm.ID_VidSF <> 6 Then
            reader = New DataTableReader(dat)
            If Not reader.Read Then
                MsgBox("СФ не содержит спецификации.")
                Exit Function
            End If
            Do
                xlSheet.Rows(rowTempl).Copy()
                InsertRow(row)
                'row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
                'ClearAllCells(row)
                ShiftRow(row, 2)
                row.ClearContents()
                row.Font.Bold = 0
                FillRow(row, "", Shift)
                ShiftRow(row, 2)
                'If bw.CancellationPending Then
                '    e.Cancel = True
                '    CancelReport()
                '    Return False
                'End If
            Loop Until Not reader.Read
            NomPP += 1
        ElseIf frm.ID_VidSF = 3 Or frm.ID_VidSF = 6 Then 'ДМС или Стоматология
            IsNomPP = False
            Dim sql2 As String
            If frm.ID_VidSF = 3 Then
                sql2 = "select * from Q_SumReportDMSLimit WHERE ID_SF=" + ID_SF.ToString
            ElseIf frm.ID_VidSF = 6 Then
                sql2 = "select * from Q_SumReportStomLimit WHERE ID_SF=" + ID_SF.ToString
            End If
            Debug.Print(sql2)
            Dim command3 As New OleDbCommand(sql2, Connection)
            Dim readerSum As OleDbDataReader = command3.ExecuteReader()
            If Not readerSum.Read Then
                MsgBox("СФ не содержит спецификации.")
                Exit Function
            End If
            Do
                If frm.ID_VidSF = 3 Then
                    sql2 = "select * from Q_ReportDMS where ID_SF=" + ID_SF.ToString + " AND ФИО_пациента='" + readerSum("ФИО_пациента") + "'"
                ElseIf frm.ID_VidSF = 6 Then
                    sql2 = "select * from Q_ReportStom where ID_SF=" + ID_SF.ToString + " AND ФИО_пациента='" + readerSum("ФИО_пациента") + "'"
                End If
                Dim command4 As New OleDbCommand(sql2, Connection)
                'command4.Parameters.Add("Дата", OleDbType.DBDate, 10)
                'command4.Parameters("Дата").Value = readerSum("Дата_посещения")
                Dim reader2 As OleDbDataReader = command4.ExecuteReader()
                Dim FirstRowValue(2) As String
                Dim Col(2) As Integer
                Dim k As Integer
                RowNomPP = row.Row
                i = 0
                Col(1) = ColumnInReport("Дата_посещения", "Реестр ДМС.xlt")
                Col(2) = ColumnInReport("Диагноз_МКБ", "Реестр ДМС.xlt")
                Do While reader2.Read
                    xlSheet.Rows(rowTempl).Copy()
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.Font.Bold = 0
                    FillRowOleDB(row, reader2, "", Shift + 1)
                    xlSheet.Cells(row.Row, 1 + Shift).Value = ""
                    If i > 0 Then ClearCells(row, reader2, "ФИО_пациента,Номер_полиса,Дата_рождения")
                    For k = 1 To 2
                        If i = 0 Then
                            FirstRowValue(k) = xlSheet.Cells(row.Row, Col(k)).Value
                        Else
                            If FirstRowValue(k) = xlSheet.Cells(row.Row, Col(k)).Value Then
                                xlSheet.Cells(row.Row, Col(k)) = ""
                            Else
                                FirstRowValue(k) = xlSheet.Cells(row.Row, Col(k)).Value
                            End If
                        End If
                    Next k
                    i += 1
                    ShiftRow(row, 2)
                Loop
                xlSheet.Cells(RowNomPP, 1 + Shift).Value = NomPP
                If readerSum("Колво_процедур") > 1 Or readerSum("Сумма") > readerSum("Сумма_к_оплате") Then
                    xlSheet.Rows(rowTempl).Copy()
                    ShiftRow(row, 2)
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.ClearContents()
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.ClearContents()
                    For i = 0 To reader2.FieldCount - 1
                        If reader2.GetName(i).ToLower = "услуга" Then
                            j = ColumnInReport(reader2.GetName(i), Templ)
                            xlSheet.Cells(row.Row, j).Font.Bold = True
                            xlSheet.Cells(row.Row, j).Value = "Итого по п. " + NomPP.ToString
                        End If
                        If reader2.GetName(i).ToLower = "сумма" Then
                            j = ColumnInReport(reader2.GetName(i), Templ)
                            xlSheet.Cells(row.Row, j).Font.Bold = True
                            xlSheet.Cells(row.Row, j).Value = readerSum("Сумма_к_оплате").ToString
                        End If
                    Next i
                    xlSheet.Rows(rowTempl).Copy()
                    ShiftRow(row, 2)
                    InsertRow(row)
                    ShiftRow(row, 2)
                    row.ClearContents()
                    ShiftRow(row, 2)
                End If
                NomPP += 1
                reader2.Close()
                command4.Dispose()
            Loop Until Not readerSum.Read
        End If
        xlSheet.Range("RowDelete").Delete(Excel.XlDirection.xlUp)
        xlSheet.Range("RowDelete2").Delete(Excel.XlDirection.xlUp)
        row.Delete()
        If sql <> "" Then
            Dim command As New OleDbCommand(sql, Connection)
            Dim reader2 As OleDbDataReader = command.ExecuteReader()
            Dim s1 As Char
            If reader2.Read Then
                Summa = reader2(0).ToString
                xlSheet.Range("Сумма").Cells(1, 1).Value = Summa
                s = GetSummaString(Int(Summa), 1, "рубль", "рубля", "рублей") + " " + GetSummaString(Int(Summa * 100) - Int(Summa) * 100, 2, "копейка", "копейки", "копеек")
                s1 = UCase(s(0))
                s = s1 + Right(s, Len(s) - 1)
                s = "Всего по реестру: " + xlSheet.Range("Сумма").Cells(1, 1).Text + " руб. (" + s + ")"
                If frm.ID_VidSF <> 4 And frm.ID_VidSF <> 5 Then xlSheet.Range("СуммаПрописью").Cells(1, 1).Value = s
            End If
        End If
        If Customer <> "" Then xlSheet.Range("Заказчик").Cells(1, 1).Value = Customer
        If frm.ID_VidSF > 2 And frm.ID_VidSF < 6 Then
            Dim row_, col As Integer
            Dim QuerySumP As String
            If frm.ID_VidSF = 3 Then
                QuerySumP = "Q_SumDMSByPodr"
                row_ = xlSheet.Range("Заказчик").Row + 6
                col = xlSheet.Range("Заказчик").Column - 1
            ElseIf frm.ID_VidSF = 4 Then
                QuerySumP = "Q_SumBeznalByPodr"
                row_ = xlSheet.Range("Сумма").Row + 7
                col = 2
            ElseIf frm.ID_VidSF = 5 Then
                QuerySumP = "Q_SumBeznalByPodr"
                row_ = xlSheet.Range("Сумма").Row + 12
                col = 2
            End If
            Dim command5 As New OleDbCommand("select * from " + QuerySumP + " WHERE ID_SF=" + ID_SF.ToString, Connection)
            Dim reader3 As OleDbDataReader = command5.ExecuteReader()
            Do While reader3.Read
                xlSheet.UsedRange.Cells(row_, col).Value = reader3("Код_подразделения").ToString + "-"
                xlSheet.UsedRange.Cells(row_, col + 1).Value = reader3("Сумма").ToString
                row_ += 1
            Loop
        End If
        If Header <> "" Then xlSheet.Range("HReport").Cells(1, 1).Value = Header
        If Header2 <> "" Then xlSheet.Range("HReport2").Cells(1, 1).Value = Header2

        Return True
Err_h:
        ErrMess(Err.Description)
        If Not (xlBook Is Nothing) Then xlBook.Close(0)
        Return False
    End Function

    Public Function ColumnInReport(ByVal ColInQuery As String, ByVal RepName As String) As Integer
        If RepName = "Реестр ДМС.xlt" Then
            Select Case ColInQuery
                Case "Дата_посещения"
                    Return 5
                Case "Номер_полиса"
                    Return 6
                Case "Диагноз_МКБ"
                    Return 7
                Case "ФИО_пациента"
                    Return 3
                Case "Дата_рождения"
                    Return 4
                Case "Услуга"
                    Return 9
                Case "Сумма"
                    Return 12
            End Select
        ElseIf RepName = "Реестр стоматология.xlt" Then
            Select Case ColInQuery
                Case "Дата_посещения"
                    Return 5
                Case "Номер_полиса"
                    Return 6
                Case "ФИО_пациента"
                    Return 3
                Case "Дата_рождения"
                    Return 4
                Case "Услуга"
                    Return 11
                Case "Сумма"
                    Return 14
            End Select
        End If

    End Function

    Public Function Rep_IncomesByPodrAndStat(ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim s As String = ""
        Dim Query2 As String
        Dim row, colStat As Excel.Range
        Dim rowTempl, colIndex, rowIndex, ID_Podr, col, i As Integer
        Dim dTemp As Single
        'Dim b As Boolean
        Dim dtStat As DataTable
        Dim frm As New frmRepInByPodrAndStat
        On Error GoTo Err_h

        If Connection.State = ConnectionState.Closed Then Connection.Open()
        'If frm.ShowDialog() = DialogResult.Cancel Then
        '    Return False
        'End If
        Select Case ID_VidReport
            Case 0
                Query = "Q_IncomesByPodrAndStatNal"
            Case 1
                Query = "Q_IncomesByPodrAndStatBeznal"
            Case 2
                Query = "Q_IncomesByPodrAndStatSold"
            Case 3
                Query = "Q_IncomesByPodrAndStatDMS"
            Case 4
                Query = "Q_IncomesByPodrAndStatUMO"
            Case 5
                Query = "Q_IncomesByPodrAndStat"
        End Select
        IsNomPP = False
        If Not GetExcelSheet() Then
            Return False
        End If
        If IsDebug Then
            objExcel.Visible = True
        End If
        Header = " " + sPeriod.ToLower
        s = xlSheet.Range("HReport").Cells(1, 1).Value()
        s += Header
        xlSheet.Range("HReport").Cells(1, 1).Value = s
        xlSheet.Range("Source").Cells(1, 1).Value = ID_VidReport.ToString
        rowTempl = xlSheet.Range("Row").Row
        row = xlSheet.Range("Row")
        'sFilterPeriod = GetPeriodsFilter(frm.PerFrom, frm.PerTo)
        'If sFilterPeriod = "" Then Return False
        Query2 = "select ID_Podr,Код_подразделения,Подразделение,ID_StatCalc,SUM(Q.Сумма) AS Сумма FROM " + Query + " AS Q WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_Podr,Код_подразделения,Подразделение,ID_StatCalc"
        Debug.Print(Query2)
        Dim command As New OleDbCommand(Query2, Connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        If Not reader.Read Then Return False
        Do While (1 = 1)
            xlSheet.Rows(rowTempl).Copy()
            Debug.Print(row.Row)
            InsertRow(row)
            ShiftRow(row, 0)
            row.Font.Bold = 0
            ClearAllCells(row)
            'row.NumberFormat = "0,00"
            row.Cells(1, 1) = reader("Подразделение").ToString
            col = 3
            ID_Podr = reader("ID_Podr")
            Do While (ID_Podr = reader("ID_Podr"))
                dTemp = reader("Сумма")
                row.Cells(1, col) = dTemp
                Debug.Print(reader("ID_StatCalc").ToString)
                Debug.Print(dTemp)
                col += 1
                If col = 9 Then col += 1
                If Not reader.Read Then
                    ShiftRow(row, 1)
                    rowIndex = row.Row
                    row.Delete()
                    's = "=СУММ(R[-15]C:R[-1]C)"
                    i = rowIndex - 7
                    s = "=СУММ(R[-" + i.ToString + "]C:R[-1]C)"
                    For col = 2 To 11
                        xlSheet.Cells(rowIndex, col).FormulaR1C1Local = s
                        'xlSheet.Cells(rowIndex, col).FormulaR1C1 = s
                    Next col
                    Return True
                End If
                If bw.CancellationPending Then
                    e.Cancel = True
                    CancelReport()
                    Return False
                End If
            Loop
            ShiftRow(row, 1)
        Loop
        row.Delete()
        Return True
        Exit Function
Err_h:
        ErrMess(Err.Description)
        Return False
    End Function

    Public Function Rep_IncomesByStat() As Boolean
        Dim sum As Decimal
        Dim s As String = ""
        Dim sPeriod As String = ""
        Dim Query As String
        Dim row, colStat As Excel.Range
        Dim rowTempl, colIndex, rowIndex, ID_PerFrom, ID_PerTo, ID_Podr, col, ID_VidReport, i As Integer
        'Dim b As Boolean
        Dim dtStat As DataTable
        'Dim frm As New frmFilterPeriods
        On Error GoTo Err_h

        If Connection.State = ConnectionState.Closed Then Connection.Open()
        IsNomPP = False
        If Not GetExcelSheet() Then
            Return False
        End If
        If IsDebug Then
            objExcel.Visible = True
        End If
        Header = " " + sPeriod.ToLower
        s = xlSheet.Range("HReport").Cells(1, 1).Value()
        s += Header
        xlSheet.Range("HReport").Cells(1, 1).Value = s
        row = xlSheet.Range("Row")
        Query = "Q_incomesByStat"
        'sFilterPeriod = GetPeriodsFilter(frm.PerFrom, frm.PerTo)
        'If sFilterPeriod = "" Then Return False
        Dim command As New OleDbCommand("SELECT ID_StatCalc, ID_VidMedStrah, SUM(Q.Сумма) AS Сумма FROM " + Query + " AS Q WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_StatCalc, ID_VidMedStrah", Connection)
        Debug.Print("SELECT ID_StatCalc, ID_VidMedStrah, SUM(Q.Сумма) AS Сумма FROM " + Query + " AS Q WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_StatCalc, ID_VidMedStrah")
        Dim reader As OleDbDataReader = command.ExecuteReader()
        col = 1
        Do While reader.Read
            row.Cells(1, col) = reader("Сумма").ToString
            col += 1
        Loop
        'command = Nothing
        reader.Close()
        reader = Nothing
        row = xlSheet.Range("RowAvans")
        Query = "Q_incomesByStatAvans"
        Dim command2 As New OleDbCommand("SELECT ID_StatCalc, SUM(Q.Сумма) AS Сумма FROM " + Query + " As Q WHERE ID_Period IN (" + sFilterPeriod + ") GROUP BY ID_StatCalc", Connection)
        Dim reader2 As OleDbDataReader = command2.ExecuteReader()
        col = 1
        Do While reader2.Read
            row.Cells(1, col) = reader2("Сумма").ToString
            col += 1
        Loop
        Return True
        Exit Function
Err_h:
        ErrMess(Err.Description)
        Return False
    End Function

    Public Sub ShiftRow(ByRef row As Excel.Range, ByVal Shift As Integer)
        If row.Columns.Count = 256 Then
            row = xlSheet.Rows(row.Row - 1 + Shift)
        Else
            row = xlSheet.UsedRange.Rows(row.Row - 1 + Shift)
        End If
    End Sub

    Public Sub ShiftColumn(ByRef col As Excel.Range, ByVal Shift As Integer)
        col = xlSheet.UsedRange.Columns(col.Column - 1 + Shift)
    End Sub

    Private Sub InsertRow(ByRef row As Excel.Range)
        row.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
        If row.Columns.Count = 256 Then
            row = xlSheet.Rows(row.Row - 2)
        Else
            row = xlSheet.UsedRange.Rows(row.Row - 2)
        End If
    End Sub

    Private Sub InsertColumn(ByRef col As Excel.Range)
        'On Error Resume Next

        col.Insert(Excel.XlInsertShiftDirection.xlShiftToRight)
        col = xlSheet.UsedRange.Columns(col.Column - 1)
    End Sub

    Function ToStr(ByVal ColName As String) As String
        If IsDBNull(reader(ColName)) Then
            ToStr = ""
        Else
            ToStr = CStr(reader(ColName))
        End If
    End Function

    Private Sub FillRow(ByRef row As Excel.Range, Optional ByVal ExceptColumns As String = "", Optional ByVal Shift As Integer = 0, Optional ByRef red As DataTableReader = Nothing)
        If (red Is Nothing) Then red = reader
        Dim i, col As Integer
        col = 1
        If IsNomPP Then
            row.Cells(1, 1 + Shift) = NomPP
            col += 1
            NomPP += 1
        End If
        For i = 1 To reader.FieldCount
            If (ExceptColumns = "" Or Not (ExceptColumns.ToLower Like "*" & reader.GetName(i - 1).ToLower & "*")) And Not (reader.GetName(i - 1) Like "ID*") Then
                If reader(i - 1).GetType.Name = "DateTime" Then
                    row.Cells(1, col + Shift) = FormatDateTime(reader(i - 1), DateFormat.ShortDate)
                Else
                    row.Cells(1, col + Shift) = reader(i - 1).ToString
                End If
                col += 1
            End If
        Next i
        'row = xlSheet.UsedRange.Rows(row.Row)
    End Sub

    Private Sub FillRowOleDB(ByRef row As Excel.Range, ByRef reader As OleDbDataReader, Optional ByVal ExceptColumns As String = "", Optional ByVal Shift As Integer = 0)
        On Error GoTo Err_h
        Dim i, col As Integer
        col = 1
        If IsNomPP Then
            row.Cells(1, 1 + Shift) = NomPP
            col += 1
            NomPP += 1
        End If
        For i = 1 To reader.FieldCount
            If (ExceptColumns = "" Or Not (ExceptColumns.ToLower Like "*" & reader.GetName(i - 1).ToLower & "*")) And Not (reader.GetName(i - 1) Like "ID*") Then
                If reader(i - 1).GetType.Name = "DateTime" Then
                    row.Cells(1, col + Shift).Value = FormatDateTime(reader(i - 1), DateFormat.ShortDate)
                Else
                    row.Cells(1, col + Shift).Value = reader(i - 1).ToString
                End If
                col += 1
            End If
        Next i
        Exit Sub
Err_h:
        ErrMess(Err.Description, "FillRowOleDB")
        'row = xlSheet.UsedRange.Rows(row.Row)
    End Sub

    Private Sub ClearCells(ByRef row As Excel.Range, ByRef reader As OleDbDataReader, ByVal Columns As String)
        Dim i, col As Integer
        For i = 1 To reader.FieldCount
            If (Columns = "" Or Columns.ToLower Like "*" & reader.GetName(i - 1).ToLower & "*") And Not (reader.GetName(i - 1) Like "ID*") Then
                col = ColumnInReport(reader.GetName(i - 1), Templ)
                row.Cells(1, col) = ""
            End If
        Next i
    End Sub

    Private Function GetPeriodsFilter(ByRef PerFrom As Period, ByVal PerTo As Period) As String

        If PerFrom.dNach > PerTo.dOkon Then
            MsgBox("Некорректно задан интервал.")
            Return ""
        End If
        Dim command_ As New OleDbCommand("SELECT * FROM D_Periods WHERE Дата_начала>=" + FormatDateSQL(PerFrom.dNach) + " and Дата_окончания<=" + FormatDateSQL(PerTo.dOkon) + " and Месяц=True", Connection)
        Dim reader_ As OleDbDataReader = command_.ExecuteReader()
        Dim str As String = ""
        Do While reader_.Read
            str = str + reader_("ID").ToString + ","
        Loop
        If str <> "" Then
            str = str.Substring(0, str.Length - 1)
        End If
        Return str
    End Function

    Public Function CreateReportByQuery(ByVal Query As String, ByVal SenderName As String, ByRef bw As System.ComponentModel.BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs) As Boolean
        Dim RC As Boolean

        Select Case Query
            Case "Q_Preiskurant"
                If SenderName = "mnuPreiskurantBas" Then
                    '    RC = Rep_Preiskurant("Код_подразделения<>20 AND Код_услуги_по_номенклатуре<>''")
                    RC = Rep_Preiskurant("Код_услуги_по_номенклатуре<>''", bw, e)
                ElseIf SenderName = "mnuPreiskurantOther" Then
                    RC = Rep_Preiskurant("IsNull(Код_услуги_по_номенклатуре)=True", bw, e)
                End If
            Case "Q_ReestrUslug"
                RC = Rep_ReestrUslug(bw, e)
            Case "Q_IncomesByPodrAndStat"
                RC = Rep_IncomesByPodrAndStat(bw, e)
            Case "Q_Akt"
                RC = Rep_Akt(SenderName)
                If RC Then objWord.Visible = True
                objWord = Nothing
                Exit Function
            Case "Q_ReportBeznal"
                RC = Rep_ReestrSF()
            Case "Q_Zakaz"
                RC = Rep_Zakaz()
            Case "Q_IncomesByStat"
                RC = Rep_IncomesByStat()
            Case "Q_UslugiByStat"
                RC = Rep_ReestrUslugByStat(bw, e)
        End Select
        Return RC
    End Function

    Public Sub SaveReport()
        Dim frmSave As New frmSaveReport
        Dim dlgSave As New SaveFileDialog()

        If objExcel.Visible = False Then
            If frmSave.ShowDialog(frmMDI) = System.Windows.Forms.DialogResult.OK Then
                If frmSave.optOpen.Checked Then
                    If Not objExcel.Visible Then
                        objExcel.Visible = True
                    End If
                    objExcel.WindowState = Excel.XlWindowState.xlMaximized
                Else
                    With dlgSave
                        '.InitialDirectory = "c:\"
                        .Filter = "Файлы MS Excel (*.xls)|*.xls|Все файлы (*.*)|*.*"
                        .FilterIndex = 1
                        .RestoreDirectory = True
                    End With
                    If dlgSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                        xlBook.SaveAs(dlgSave.FileName)
                    End If
                End If
            End If
        End If
        objExcel = Nothing
        Exit Sub
Err_h:
        ErrMess(Err.Description, "CreateReport")
    End Sub

    Private Function GetParamForReport(ByVal Query As String) As Boolean
        Dim b As Boolean
        Dim frmRepReestr As New frmRepReestrByStat
        Dim frmRepIn As New frmRepInByPodrAndStat
        Dim frmFilterP As New frmFilterPeriods
        Dim frmFilterZ As New frmFilterZakaz
        Dim frmRepByStat As New frmRepReestrByStat
        Dim frmFilterD As New frmFilterData
        Dim frm2 As frmList
        Dim SelRow As Integer
        b = False
        Select Case Query
            Case "Q_Preiskurant"
                b = (frmFilterDataSimple.ShowDialog() = DialogResult.Cancel)
            Case "Q_ReestrUslug"
                frmFilterD.optNoPay.Visible = False
                frmFilterD.optAll.Location = frmFilterD.optNoPay.Location
                b = (frmFilterD.ShowDialog() = DialogResult.Cancel)

            Case "Q_Zakaz"
                frm2 = frmMDI.ActiveMdiChild
                If frm2 Is Nothing Then
                    MsgBox("Необходимо открыть список счёт-фактур и выбрать нужный документ.")
                    Return False
                End If
                If frm2.dgwList.SelectedCells.Count = 0 Then
                    MsgBox("Выберите счет-фактуру.")
                    Return False
                End If
                SelRow = frm2.dgwList.SelectedCells(0).RowIndex
                ID_SF = frm2.dgwList.Item("ID", SelRow).Value
                frmFilterZ.ID_Sf = ID_SF
                b = (frmFilterZ.ShowDialog() = DialogResult.Cancel)
                If b = True Then Return False
                sFIO = frmFilterZ.cmbFIO.Text

            Case "Q_IncomesByStat"
                b = (frmFilterP.ShowDialog() = DialogResult.Cancel)
                sFilterPeriod = GetPeriodsFilter(frmFilterP.PerFrom, frmFilterP.PerTo)
                ID_PerFrom = frmFilterP.cbPerFrom.SelectedValue
                ID_PerTo = frmFilterP.cbPerTo.SelectedValue
                If sFilterPeriod = "" Then b = True
                If ID_PerTo <> ID_PerFrom Then
                    sPeriod = frmFilterP.cbPerFrom.Text + "-" + frmFilterP.cbPerTo.Text
                Else
                    sPeriod = frmFilterP.cbPerFrom.Text
                End If

            Case "Q_UslugiByStat"
                b = (frmRepByStat.ShowDialog() = DialogResult.Cancel)
                sFilterPeriod = GetPeriodsFilter(frmRepByStat.PerFrom, frmRepByStat.PerTo)
                If sFilterPeriod = "" Then b = True
                ID_VidReport = frmRepByStat.cbVidReport.SelectedIndex

            Case “Q_SumByStat”
                If frmRepByStat.PerTo.ID <> frmRepByStat.PerFrom.ID Then
                    sTitle = frmRepByStat.cbPerFrom.Text + "-" + frmRepByStat.cbPerTo.Text
                Else
                    sTitle = frmRepByStat.cbPerFrom.Text
                End If
                ID_VidReport = frmRepByStat.cbVidReport.SelectedIndex
                If ID_VidReport = -1 Then
                    MsgBox("Необходимо выбрать вид отчета.")
                Else
                    sTitle = sTitle + " (" + frmRepByStat.cbVidReport.Text + ")"
                End If
                If frmRepByStat.optPodr.Checked Then
                    ID_Podr = frmRepByStat.cbPodr.SelectedValue
                Else
                    ID_Podr = -1
                End If
            Case Else
                If Len(Query) >= Len("Q_IncomesByPodrAndStat") Then
                    If Query.Substring(0, Len("Q_IncomesByPodrAndStat")) = "Q_IncomesByPodrAndStat" Then
                        b = (frmRepIn.ShowDialog() = DialogResult.Cancel)
                        ID_PerFrom = frmRepIn.cbPerFrom.SelectedValue
                        ID_PerTo = frmRepIn.cbPerTo.SelectedValue
                        ID_VidReport = frmRepIn.cbVidReport.SelectedIndex
                        sFilterPeriod = GetPeriodsFilter(frmRepIn.PerFrom, frmRepIn.PerTo)
                        If ID_PerTo <> ID_PerFrom Then
                            sPeriod = frmRepIn.cbPerFrom.Text + "-" + frmRepIn.cbPerTo.Text
                        Else
                            sPeriod = frmRepIn.cbPerFrom.Text
                        End If
                        If sFilterPeriod = "" Then b = True
                    End If
                End If

        End Select


        Return (Not b)

    End Function

    Private Sub CancelReport()
        If objExcel.Visible = False Then
            xlBook.Close(SaveChanges:=False)
            objExcel.DisplayAlerts = False
            objExcel.Quit()
        End If
        xlSheet = Nothing
        xlBook = Nothing
        objExcel = Nothing
        System.GC.Collect()
    End Sub

End Module
