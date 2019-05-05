Imports System.Data.OleDb
Imports System.Windows.Forms

Public Module modBas
    Public PathDB As String
    Public Connection As OleDbConnection
    'Public ADOConn As ADODB.Connection

    Public Structure Period
        Private dNull As Date
        Private mID As Integer
        Public dNach As Date
        Public dOkon As Date
        Private mName As String

        Friend ReadOnly Property Name() As String
            Get
                If mID > 0 Then
                    Return mName
                Else
                    dNull = New Date(1900, 1, 1)
                    If dNach > dNull And dEnd > dNull Then
                        mName = "c " + FormatDateTime(dNach, DateFormat.ShortDate) + " по " + FormatDateTime(dOkon, DateFormat.ShortDate)
                    End If
                    Return mName
                End If
            End Get
        End Property

        Friend Property ID() As Integer
            Get
                Return mID
            End Get
            Set(ByVal value As Integer)
                mID = value
                Dim sql As String
                sql = "SELECT * FROM D_Periods WHERE ID=" + value.ToString
                Debug.Print(sql)
                If Connection.State = ConnectionState.Closed Then Connection.Open()
                Dim command As New OleDbCommand(sql, Connection)
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read Then
                    dNach = reader("Дата_начала")
                    dOkon = reader("Дата_окончания")
                    mName = reader("Наименование")
                Else
                    dNach = New Date(1900, 1, 1)
                    dOkon = New Date(1900, 1, 1)
                    mID = 0
                    mName = ""
                End If
            End Set
        End Property

        Friend Sub SetPeriodByDate(ByVal d As Date, Optional ByVal Vid As String = "m")
            Dim sql As String
            sql = "SELECT * FROM D_Periods WHERE Дата_начала>=? AND Дата_окончания<=? AND Vid='" + Vid + "'"
            Debug.Print(sql)
            Dim Param1 As OleDbParameter
            Dim Param2 As OleDbParameter
            Dim command As New OleDbCommand(sql, _
                Connection)
            Dim adap As New OleDbDataAdapter()
            'If Today.Date > "06/06/2009" Then
            '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
            '    End
            'End If
            Param1 = command.Parameters.Add("Param1", OleDbType.DBDate)
            Param1.Value = d
            Param2 = command.Parameters.Add("Param2", OleDbType.DBDate)
            Param2.Value = d

            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.Read Then
                mID = reader("ID")
                dNach = reader("Дата_начала")
                dOkon = reader("Дата_окончания")
                mName = reader("Наименование")
            Else
                mID = 0
                dNach = New Date(1900, 1, 1)
                dOkon = New Date(1900, 1, 1)
                mName = ""
            End If
        End Sub

    End Structure

    Public Function readerBySQL(ByVal SQL As String) As OleDbDataReader
        Dim command As New OleDbCommand(SQL, Connection)
        readerBySQL = command.ExecuteReader
    End Function

    Public Function GetEndPeriod(ByVal d As Date) As Date
        Dim dEnd As Date
        Dim sql As String
        Dim s As String

        s = FormatDateSQL(d)
        sql = "SELECT * FROM D_Periods WHERE Дата_начала<=" + s + " AND Дата_окончания>=" + s
        Debug.Print(sql)
        Dim command As New OleDbCommand(sql, Connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        If reader.Read Then
            dEnd = reader("Дата_окончания")
        Else
            dEnd = d
        End If
        Return dEnd
    End Function

    'Public Overloads Function GetPeriod(ByVal d As Date) As Period
    '    Dim Per As Period
    '    Dim sql As String
    '    Dim s As String

    '    s = FormatDateSQL(d)
    '    sql = "SELECT * FROM D_Periods WHERE Дата_начала<=" + s + " AND Дата_окончания>=" + s
    '    Debug.Print(sql)
    '    Dim command As New OleDbCommand(sql, Connection)
    '    Dim reader As OleDbDataReader = command.ExecuteReader()
    '    If reader.Read Then
    '        Per.dNach = reader("Дата_начала")
    '        Per.dOkon = reader("Дата_окончания")
    '        Per.ID = reader("ID")
    '    End If
    '    Return Per
    'End Function

    Public Function GetPeriod(ByVal ID As Integer) As Period
        Dim Per As Period
        Dim sql As String

        sql = "SELECT * FROM D_Periods WHERE ID=" + ID.ToString
        Debug.Print(sql)
        Dim command As New OleDbCommand(sql, Connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        If reader.Read Then
            Per.dNach = reader("Дата_начала")
            Per.dOkon = reader("Дата_окончания")
            Per.ID = reader("ID")
        End If
        Return Per
    End Function

    Public Function GetTableToUpdate(ByVal TabName As String) As String
        Dim sql, s As String

        sql = "SELECT * FROM M_Tables WHERE Table='" + TabName + "'"
        Dim command As New OleDbCommand(sql, Connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        If reader.Read Then
            If Not IsDBNull(reader("TableToUpdate")) Then
                Return reader("TableToUpdate").ToString
            End If
        End If
        s = TabName.Replace("Q_", "R_")
        Return s
    End Function

    Public Function FormatDateSQL(ByVal d As Date) As String
        Dim s As String
        s = Format(d, "#M/d/yyyy#")
        s = s.Replace(".", "/")
        Return s
    End Function

    Public Function ConnectionString() As String
        On Error GoTo Err_h

        Dim builder As New OleDbConnectionStringBuilder
        Dim cs As String
        cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""****"";Persist Security Info=True;Mode=ReadWrite"
        cs = cs.Replace("****", PathDB)
        Debug.Print(cs)
        'With builder
        '    .Provider = "Microsoft.ACE.OLEDB.12.0"
        '    .DataSource = PathDB
        '    .Add("Mode", "ReadWrite")
        '    '   .PersistSecurityInfo = True
        'End With
        'ConnectionString = builder.ConnectionString
        ConnectionString = cs
        Exit Function
Err_h:
        'MsgBox(Err.Description)
    End Function

    '    Public Sub OpenADOConnection()
    '        On Error GoTo Err_h
    '        Dim ConnStr1 As String
    '        ConnStr1 = "Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
    '        'ConnStr1 = ""
    '        ADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Password="""";Persist Security Info=False;Data Source=" & PathDB & ";Mode=ReadWrite;" & ConnStr1
    '        ADOConn.Open()
    '        Exit Sub
    'Err_h:
    '        MsgBox("Ошибка при открытии базы данных:" & Err.Description)
    '    End Sub

    Public Function Populate(ByVal sqlString As String) As DataTable
        On Error GoTo Err_h
        Dim command As New OleDbCommand(sqlString, _
            Connection)
        Dim adapter As New OleDbDataAdapter()
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        Debug.Print(sqlString)
        adapter.SelectCommand = command
        Dim table As New DataTable()
        table.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adapter.MissingMappingAction = MissingMappingAction.Ignore

        adapter.Fill(table)
        Return table
        Exit Function
Err_h:
        ErrMess(Err.Description)

    End Function

    Public Function PopulateInterval(ByVal sqlString As String, ByVal dBegin As Date, ByVal dEnd As Date) As DataTable
        On Error GoTo Err_h
        Dim Param1 As OleDbParameter
        Dim Param2 As OleDbParameter
        Dim command As New OleDbCommand(sqlString, _
            Connection)
        Dim adap As New OleDbDataAdapter()
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        adap.SelectCommand = command
        Param1 = adap.SelectCommand.Parameters.Add("Param1", OleDbType.DBDate)
        Param1.Value = dBegin

        Param2 = adap.SelectCommand.Parameters.Add("Param2", OleDbType.DBDate)
        Param2.Value = dEnd

        Dim table As New DataTable()
        table.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adapter.MissingMappingAction = MissingMappingAction.Ignore

        adap.Fill(table)
        Return table
        Exit Function
Err_h:
        ErrMess(Err.Description)

    End Function

    Public Function IsColumnInDgw(ByRef dgw As DataGridView, ByVal ColName As String) As Boolean
        'Dim i As Integer
        'On Error Resume Next
        'If ColName Is Nothing Then Return False
        'i = dgw.Columns(ColName).Index
        'Return Err.Number = 0
        Return dgw.Columns.Contains(ColName)
    End Function

    Sub ErrMess(ByVal str As String, Optional ByVal Source As String = "")
        If Source <> "" Then
            str = "Ошибка в " + Source + ":" + str
        End If
        If frmMDI.mnuDebug.Checked Then
            MsgBox(str)
        Else
            Debug.Print(str)
        End If
    End Sub

    Public Function IsKeyInCollection(ByVal Col As Collection, ByVal Name As String) As Boolean
        Return Col.Contains(Name)
    End Function

    Public Function IsKeyInCollectionOld(ByVal Col As Collection, ByVal Name As String) As Boolean
        Dim s As String
        On Error Resume Next
        s = Col(Name).ToString()

        Return Err.Number = 0
    End Function

    Public Function IsColumnInRow(ByRef dRow As Data.DataRow, ByVal Name As String)
        'Dim s As String
        'On Error Resume Next
        's = dRow(Name).ToString()

        'Return Err.Number = 0
        Return dRow.Table.Columns.Contains(Name)
    End Function

    Public Function IsColumnInTable(ByRef dTable As Data.DataTable, ByVal Name As String)
        'Dim s As String
        'On Error Resume Next
        's = dTable.Columns(Name).ColumnName

        'Return Err.Number = 0
        Return dTable.Columns.Contains(Name)
    End Function

    Public Sub DeselectAllCells(ByRef frm As frmList)

        DeselectAllCellsInDgw(frm.dgwList)
        'If frm.dgwList.Rows.Count > 0 Then
        '    frm.dgwList.Rows(0).Cells(frm.dgwList.Columns.Count - 1).Selected = True
        'End If
        DeselectAllCellsInDgw(frm.dgwSumma)
        DeselectAllCellsInDgw(frm.dgwList2)
        DeselectAllCellsInDgw(frm.dgwSumma2)

    End Sub

    Public Sub DeselectAllCellsInDgw(ByRef dgw As DataGridView)
        'Dim i, row, col As Integer

        If dgw.Visible = False Or dgw.SelectedCells.Count = 0 Then Exit Sub
        Do While dgw.SelectedCells.Count > 0
            dgw.SelectedCells(0).Selected = False
        Loop
    End Sub

    Public Sub MakeCopyDB()
        Dim Dest As String
        Dim DBName As String = "UchetMU.mdb"

        Dest = Left(PathDB, Len(PathDB) - Len(DBName) - 1)
        Dest += "\Copy " + FormatDateTime(Today, DateFormat.ShortDate)
        On Error Resume Next
        MkDir(Dest)
        If Err.Number <> 75 And Err.Number > 0 Then
            MsgBox("Отсутсутствуют права на запись резервное копии. Обратитесь к администратору.")
            Exit Sub
        End If
        Dest += "\" + DBName
        Debug.Print(Dest)
        FileCopy(PathDB, Dest)
    End Sub

    Public Sub UpdateSQL(ByVal sql As String)
        Dim adap As New OleDbDataAdapter

        adap.UpdateCommand = New OleDbCommand(sql, Connection)
        adap.UpdateCommand.ExecuteNonQuery()

    End Sub

    Public Sub ReFillDependsCombo(ByRef frm As frmList, ByVal NomTable As Integer, ByVal col As Integer, ByVal row As Integer)
        Dim i, rowBegin, rowEnd As Integer
        Dim dgw As DataGridView
        Dim ds As New DataTable
        Dim ad As OleDbDataAdapter
        Dim ColName, CellName As String
        Dim Cell As DataGridViewComboBoxCell
        On Error GoTo Err_h

        'If Not frm.IsLinkingWithAnother Then Exit Sub
        dgw = frm.GetdgwByNom(NomTable)
        If row = -1 Then
            rowBegin = 0
            rowEnd = dgw.Rows.Count - 1
        Else
            rowBegin = row
            rowEnd = row
        End If
        ColName = dgw.Columns(col).Name
        If NomTable = 1 And (frm.TableName = "R_Uslugi" Or frm.TableName = "R_Calc") And ColName = "ID_Podr" Then
            For row = rowBegin To rowEnd
                Cell = dgw.Item("Услуга", row)
                If IsDBNull(dgw.Item(ColName, row).Value) Then Exit Sub
                i = dgw.Item(ColName, row).Value
                ds.Locale = System.Globalization.CultureInfo.InvariantCulture
                'adaptercombo = Nothing
                ad = New OleDbDataAdapter("select * from Q_MedUslug_NoCalc Where ID_Podr=" + i.ToString + " ORDER BY Наименование", Connection)
                ds = New DataTable
                ad.Fill(ds)
                CellName = frm.TableName + ".ID_Podr=" + i.ToString
                'DSCombos.Add(ds, CellName)
                'ComboAdapters.Add(ad, CellName)
                If IsKeyInCollection(frm.DSCombos, CellName) Then
                    frm.DSCombos.Remove(CellName)
                    frm.ComboAdapters.Remove(CellName)
                    frm.Combos.Remove(CellName)
                End If
                If ds.Rows.Count = 0 Then
                    Cell.Value = Nothing
                    Cell.ReadOnly = True
                    Exit Sub
                Else
                    Cell.ReadOnly = False
                End If
                frm.DSCombos.Add(ds, CellName)
                frm.ComboAdapters.Add(ad, CellName)
                frm.Combos.Add(Cell, CellName)

                Cell.DataSource = frm.DSCombos(CellName)
                Cell.ValueMember = "ID"
                Cell.DisplayMember = "Наименование"
                Debug.Print(ds.Rows.Count)
            Next row

            'Cell.Name = CellName
            'Cell.Add(combo, sTable)

        End If
        Exit Sub
Err_h:
        ErrMess(Err.Description, "modMedUslug.ReFillDependsCombo")
    End Sub

    Public Function GetOLEDBType(ByVal sType As String) As OleDbType
        Select Case sType
            Case "DateTime"
                Return OleDbType.DBDate
            Case "Int32"
                Return OleDbType.BigInt
            Case "Single"
                Return OleDbType.Single
            Case Else
                Return OleDbType.VarChar

        End Select
    End Function

    Public Function GetOLEDBSize(ByVal sType As String) As Integer
        Select Case sType
            Case "String"
                Return 255
            Case Else
                Return 10

        End Select
    End Function

    Public Function GetSummaMoneyString(ByVal Summa As Decimal) As String
        Dim str As String = ""
        str = GetSummaString(Int(Summa), 1, "рубль", "рубля", "рублей") + " " + GetSummaString(Int(Summa * 100) - Int(Summa) * 100, 2, "копейка", "копейки", "копеек")
        Return str
    End Function

    Public Function GetSummaString(ByVal Source As Long, ByVal Rod%, ByVal w1$, ByVal w2to4$, ByVal w5to10$) As String
        Dim str As String = ""
        SummaString(str, Source, Rod%, w1$, w2to4$, w5to10$)
        Return str
    End Function

    Public Sub SummaString(ByRef Summa$, ByVal Source As Long, ByVal Rod%, ByVal w1$, ByVal w2to4$, ByVal w5to10$)
        ' "Сумма прописью":
        '  преобразование числа из цифрого вида в символьное
        ' ==================================================
        ' Исходные данные:
        '  Source - число от 0 до 2147483647 (2^31-1)
        ' Eсли нужно оперировать с числами > 2 147 483 647
        ' замените описание переменных Source и TempValue на "AS DOUBLE"
        '
        '    далее нужно задать информацию о единице изменения
        '  Rod%   = 1 - мужской, = 2 - женский, = 3 - средний
        '     название единицы изменения:
        '  w1$     - именительный падеж единственное число (= 1)
        '  w2to4$  - родительный падеж единственное число (= 2-4)
        '  w5to10$ - родительный падеж множественное число ( = 5-10)
        '
        '  Rod% должен быть задано обязательно, название единицы может быть
        '       не задано = ""
        ' ———————————————-
        ' Результат: Summa$ - запись прописью
        '
        '================================
        Dim TempValue As Long
        '
        If Source& = 0 Then
            Summa$ = RTrim$("ноль " + w5to10$) : Exit Sub
        End If
        '
        TempValue = Source : Summa$ = ""
        ' единицы
        Call SummaStringThree(Summa$, TempValue, Rod%, w1$, w2to4$, w5to10$)
        If TempValue = 0 Then Exit Sub
        ' тысячи
        Call SummaStringThree(Summa$, TempValue, 2, "тысяча", "тысячи", "тысяч")
        If TempValue = 0 Then Exit Sub
        ' миллионы
        Call SummaStringThree(Summa$, TempValue, 1, "миллион", "миллиона", "миллионов")
        If TempValue = 0 Then Exit Sub
        ' миллиардов
        Call SummaStringThree(Summa$, TempValue, 1, "миллиард", "миллиарда", "миллиардов")
        If TempValue = 0 Then Exit Sub
        '
        ' Eсли нужно оперировать с числами > 2 147 483 647
        ' измените тип переменных (см. выше) и добавьте эту строку для триллионов:
        ' CALL SummaStringThree(Summa$, TempValue#, 1, "трилллион","триллиона", "триллионов")
        ' IF TempValue# = 0 THEN EXIT SUB
        '
        ' Что идет после триллионов, я плохо представляю...
        '
    End Sub

    Sub SummaStringThree(ByRef Summa$, ByRef TempValue As Long, ByVal Rod%, ByVal w1$, ByVal w2to4$, ByVal w5to10$)
        '
        '  Формирования строки для трехзначного числа:
        '  (последний трех знаков TempValue
        ' Eсли нужно оперировать с числами > 2 147 483 647
        ' замените в описании на  TempValue AS DOUBLE
        '====================================
        Dim Rest%, Rest1%, EndWord$, s1$, s10$, s100$
        '
        Rest% = TempValue& Mod 1000
        TempValue& = TempValue& \ 1000
        If Rest% = 0 Then    ' последние три знака нулевые
            If Summa$ = "" Then Summa$ = w5to10$ + " "
            Exit Sub
        End If
        '
        ' начинаем подсчет с Rest
        EndWord$ = w5to10$
        ' сотни
        Select Case Rest% \ 100
            Case 0 : s100$ = ""
            Case 1 : s100$ = "сто "
            Case 2 : s100$ = "двести "
            Case 3 : s100$ = "триста "
            Case 4 : s100$ = "четыреста "
            Case 5 : s100$ = "пятьсот "
            Case 6 : s100$ = "шестьсот "
            Case 7 : s100$ = "семьсот "
            Case 8 : s100$ = "восемьсот "
            Case 9 : s100$ = "девятьсот "
        End Select
        '
        ' десятки
        Rest% = Rest% Mod 100 : Rest1% = Rest% \ 10
        s1$ = ""
        Select Case Rest1%
            Case 0 : s10$ = ""
            Case 1  ' особый случай
                Select Case Rest%
                    Case 10 : s10$ = "десять "
                    Case 11 : s10$ = "одиннадцать "
                    Case 12 : s10$ = "двенадцать "
                    Case 13 : s10$ = "тринадцать "
                    Case 14 : s10$ = "четырнадцать "
                    Case 15 : s10$ = "пятнадцать "
                    Case 16 : s10$ = "шестнадцать "
                    Case 17 : s10$ = "семнадцать "
                    Case 18 : s10$ = "восемнадцать "
                    Case 19 : s10$ = "девятнадцать "
                End Select
            Case 2 : s10$ = "двадцать "
            Case 3 : s10$ = "тридцать "
            Case 4 : s10$ = "сорок "
            Case 5 : s10$ = "пятьдесят "
            Case 6 : s10$ = "шестьдесят "
            Case 7 : s10$ = "семьдесят "
            Case 8 : s10$ = "восемьдесят "
            Case 9 : s10$ = "девяносто "
        End Select
        '
        If Rest1% <> 1 Then  ' единицы
            Select Case Rest% Mod 10
                Case 1
                    Select Case Rod%
                        Case 1 : s1$ = "один "
                        Case 2 : s1$ = "одна "
                        Case 3 : s1$ = "одно "
                    End Select
                    EndWord$ = w1$
                Case 2
                    If Rod% = 2 Then s1$ = "две " Else s1$ = "два "
                    EndWord$ = w2to4$
                Case 3 : s1$ = "три " : EndWord$ = w2to4$
                Case 4 : s1$ = "четыре " : EndWord$ = w2to4$
                Case 5 : s1$ = "пять "
                Case 6 : s1$ = "шесть "
                Case 7 : s1$ = "семь "
                Case 8 : s1$ = "восемь "
                Case 9 : s1$ = "девять "
            End Select
        End If
        '
        ' сборка строки
        Summa$ = RTrim$(RTrim$(s100$ + s10$ + s1$ + EndWord$) + " " + Summa$)
    End Sub

    Public Function IsObjectExist(ByVal Name As String) As Boolean
        Dim dt As DataTable
        Dim command As New OleDbCommand("select * from " + Name, Connection)
        Dim adapter As New OleDbDataAdapter()
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        adapter.SelectCommand = command
        Dim table As New DataTable()
        table.Locale = System.Globalization.CultureInfo.InvariantCulture
        'adapter.MissingMappingAction = MissingMappingAction.Ignore
        On Error Resume Next
        adapter.Fill(table)

        Return Err.Number = 0
    End Function

    Public Function DateToStr(ByVal d As Date) As String
        Dim str As String
        'str = "«" + d.Day.ToString + "» " + MonthNameByNom(d.Month) + " " + d.Year.ToString + " г."
        str = "«"
        str += d.Day.ToString
        str += "» " + MonthNameByNom(d.Month) + " " + d.Year.ToString + " г."
        Return str
    End Function

    Public Function MonthNameByNom(ByVal MonthNom As Integer) As String
        Select Case MonthNom
            Case 1
                Return "янаваря"
            Case 2
                Return "февраля"
            Case 3
                Return "марта"
            Case 4
                Return "апреля"
            Case 5
                Return "мая"
            Case 6
                Return "июня"
            Case 7
                Return "июля"
            Case 8
                Return "августа"
            Case 9
                Return "сентября"
            Case 10
                Return "октября"
            Case 11
                Return "ноября"
            Case 12
                Return "декабря"
            Case Else
                Return ""
        End Select

    End Function

    Public Sub UpdateValueInDgw(ByRef frm As frmList, ByVal NomTable As Integer, ByVal Row As Integer, ByVal ColName As String, ByVal Value As Object)
        Dim sql As String
        Dim dgw As DataGridView
        Dim ID As Integer
        Dim adap As New OleDbDataAdapter
        'Dim dt As DataTable
        Dim dt2 As New DataTable
        'Dim dRow, dRow2 As DataRow()


        dgw = frm.GetdgwByNom(NomTable)
        ID = dgw.Item("ID", Row).Value
        sql = "UPDATE " + frm.TableToUpdate(NomTable) + " SET " + ColName + "=? WHERE ID=" + ID.ToString
        adap.UpdateCommand = New OleDbCommand(sql, Connection)
        adap.UpdateCommand.Parameters.Add( _
                    ColName, GetOLEDBType(dgw.Columns(ColName).ValueType.Name), GetOLEDBSize(dgw.Columns(ColName).ValueType.Name))
        adap.UpdateCommand.Parameters(ColName).Value = Value
        adap.UpdateCommand.ExecuteNonQuery()

    End Sub

    Public Sub UpdateValueInDgwByID(ByRef frm As frmList, ByVal NomTable As Integer, ByVal ID As Integer, ByVal ColName As String, ByVal Value As Object)
        Dim sql As String
        Dim dgw As DataGridView
        Dim adap As New OleDbDataAdapter
        'Dim dt As DataTable
        Dim dt2 As New DataTable
        'Dim dRow, dRow2 As DataRow()


        dgw = frm.GetdgwByNom(NomTable)
        sql = "UPDATE " + frm.TableToUpdate(NomTable) + " SET " + ColName + "=? WHERE ID=" + ID.ToString
        adap.UpdateCommand = New OleDbCommand(sql, Connection)
        adap.UpdateCommand.Parameters.Add( _
                    ColName, GetOLEDBType(frm.dgwList.Columns(ColName).ValueType.Name), GetOLEDBSize(frm.dgwList.Columns(ColName).ValueType.Name))
        adap.UpdateCommand.Parameters(ColName).Value = Value
        adap.UpdateCommand.ExecuteNonQuery()

        'dt = frm.GetDataTableByNom(NomTable)
        'dRow = dt.Select("ID=" + ID.ToString)
        ''dRow(0).Delete()
        'Dim command As New OleDbCommand("SELECT * FROM " + frm.mTable(NomTable) + " WHERE ID=" + ID.ToString, Connection)
        'Dim adap2 As New OleDbDataAdapter()
        'If Today.Date > "06/06/2009" Then
        '    MsgBox("Период использования пробной версии программы истек. По вопросам приобретения обращайтесь roman_box@mail.ru")
        '    End
        'End If
        'adap2.SelectCommand = command
        'adap2.Fill(dt)
        'dRow2 = dt2.Select("ID=" + ID.ToString)
        'dRow(0) = dRow2(0)


    End Sub

    Public Sub CopyDataRow(ByRef dRowFrom As DataRow, ByRef dRowTo As DataRow)
        Dim dt As DataTable
        Dim i As Integer
        Dim ColName As String

        dt = dRowFrom.Table
        For i = 0 To dt.Columns.Count - 1
            ColName = dt.Columns(i).ColumnName
            dRowTo(ColName) = dRowFrom(ColName)
        Next i
    End Sub

    Public Function CopyTable(ByRef ConnFrom As OleDbConnection, ByRef ConnTo As OleDbConnection, ByVal sTable As String) As Boolean
        Dim dt As New DataTable
        Dim dtTo As New DataTable
        Dim dRowFrom, dRowTo As DataRow
        Dim adapFrom, adapTo As OleDbDataAdapter
        On Error GoTo Err_h

        'dt.Locale = System.Globalization.CultureInfo.InvariantCulture
        adapFrom = New OleDbDataAdapter("SELECT * FROM " + sTable, ConnFrom)
        'ComBuilderFrom = New OleDbCommandBuilder(adapFrom)
        adapFrom.Fill(dt)
        If Connection.State = ConnectionState.Closed Then Connection.Open()
        Dim DelCommand As New OleDbCommand("DELETE FROM " + sTable, Connection)
        DelCommand.ExecuteNonQuery()
        adapTo = New OleDbDataAdapter("SELECT * FROM " + sTable, Connection)
        Dim ComBuilderTo = New OleDbCommandBuilder(adapTo)
        adapTo.Fill(dtTo)
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            dRowFrom = dt.Rows(i)
            dRowTo = dtTo.NewRow
            CopyDataRow(dRowFrom, dRowTo)
            dtTo.Rows.Add(dRowTo)
            adapTo.Update(dtTo)
        Next i
        Return True
Err_h:
        Debug.Print(Err.Description)
        MsgBox("Не удалось импортировать данные. Ошибка " + Err.Description)
        Return False
    End Function

End Module
