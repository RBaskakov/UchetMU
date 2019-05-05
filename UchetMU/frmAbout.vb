Public NotInheritable Class frmAbout
    'Declare Function Message Lib "Report.dll" Alias "clsReport.Message" () As String

    Private Sub frmAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        'Dim rep As New Report.clsReport
        'rep.Message()
    End Sub

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = "О программе " + ApplicationTitle
        ' Initialize all of the text displayed on the About Box.
        ' TODO: Customize the application's assembly information in the "Application" pane of the project 
        '    properties dialog (under the "Project" menu).
        'Me.LabelProductName.Text = My.Application.Info.ProductName
        'Me.LabelVersionApp.Text = "Версия приложения " + My.Application.Info.Version.ToString
        'Me.LabelVersionDB.Text = "" ' "Версия базы данных " + GetDBVersion()
        Me.LabelCopyRight.Text = My.Application.Info.Copyright
        'Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        'Me.TextBoxDescription.Text = My.Application.Info.Description
    End Sub

    '    Private Function GetDBVersion() As String
    '        Dim Rs As New ADODB.Recordset
    '        On Error GoTo Err_h

    '        Rs.Open("select * from DBProp where Name='DBVersion'", ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If Rs.EOF Then
    '            Return ""
    '        Else
    '            Return Rs("Value").Value
    '        End If
    '        Exit Function
    'Err_h:
    '        Return ""
    '    End Function

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start("mailto:roman_box@mail.ru")
        Catch ex As Exception
            ' The error message
            'MessageBox.Show("Unable to open link that was clicked.")
        End Try

    End Sub


End Class
