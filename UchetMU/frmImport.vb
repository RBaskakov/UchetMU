Imports System.Windows.Forms
Imports System.Data.OleDb

Public Class frmImport
    Private adapFrom, adapTo As OleDbDataAdapter
    Private ComBuilderFrom, ComBuilderTo As OleDbCommandBuilder
    Private ConnFrom As OleDbConnection

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        On Error GoTo Err_h

        If txtPathDB.Text = PathDB Then
            MsgBox("Файл источника и приемника данных совпадают. Импорт невозможен.")
            Exit Sub
        End If
        ConnFrom = New OleDbConnection(ConnectionStringFrom)
        If Not CopyTable(ConnFrom, Connection, "R_SF_Stom") Then Exit Sub
        If Not CopyTable(ConnFrom, Connection, "R_Stom") Then Exit Sub
        MsgBox("Импорт успешно завершен.")
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
        Exit Sub
Err_h:
        MsgBox("Не удалось импортировать данные. Ошибка " + Err.Description)

    End Sub


    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub butCh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCh.Click
        Dim dlgOpen As New OpenFileDialog()
        Dim Dir As String
        On Error GoTo Err_h
        If txtPathDB.Text = "" Then
            Dir = "c:\"
        Else
            Dir = txtPathDB.Text
        End If

        With dlgOpen
            .InitialDirectory = Dir
            .Filter = "MS Access files (*.mdb)|*.mdb|All files (*.*)|*.*"
            .FilterIndex = 2
            .RestoreDirectory = True
        End With

        If dlgOpen.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            txtPathDB.Text = dlgOpen.FileName
        End If
        Exit Sub
Err_h:
        ErrMess(Err.Description)
    End Sub

    Public Function ConnectionStringFrom() As String
        Dim builder As New OleDbConnectionStringBuilder()
        With builder
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .DataSource = txtPathDB.Text
            .Add("Mode", "ReadWrite")
            '   .PersistSecurityInfo = True
        End With
        Return builder.ConnectionString
    End Function

End Class
