Imports System.Windows.Forms

Public Class frmDB

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        PathDB = txtPathDB.Text
        'Connect()
        SaveSetting("MedService", "Settings", "PathDB", PathDB)
        Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dlgOpen As New OpenFileDialog()
        On Error GoTo Err_h
        With dlgOpen
            .InitialDirectory = "c:\"
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

    Private Sub frmDB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtPathDB.Text = PathDB
    End Sub

    Private Sub OK_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles OK.DragOver

    End Sub
End Class
