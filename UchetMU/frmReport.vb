Public Class frmReport
    Public ReportQuery As String
    Public ReportTitle As String
    Public SenderName As String
    Private RC As Boolean

    Private bw As ComponentModel.BackgroundWorker = New ComponentModel.BackgroundWorker

    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub frmReport_Load(sender As Object, e As EventArgs) Handles Me.Load
        'lblCap.Text = "Формирование отчёта " + CStr(ReportTitle) 
        bw.WorkerReportsProgress = False
        bw.WorkerSupportsCancellation = True
        AddHandler bw.DoWork, AddressOf bw_DoWork
        AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
        'AddHandler bw.CancelAsync, AddressOf bw_CancelAsync
        If Not bw.IsBusy = True Then
            bw.RunWorkerAsync()
        End If

    End Sub
    Private Sub bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs)

        If bw.CancellationPending = True Then
            e.Cancel = True
        Else
            RC = CreateReportByQuery(ReportQuery, Me.SenderName, bw, e)
        End If
    End Sub

    Private Sub bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs)
        If e.Error IsNot Nothing Then
            MsgBox("Ошибка: " & e.Error.Message)
        End If
        Me.Visible = False
        If RC Then SaveReport()
    End Sub

    Private Sub bw_CancelAsync(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs)

        If e.Error IsNot Nothing Then
            MsgBox("Ошибка: " & e.Error.Message)
        End If
        Me.Close()

    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        If bw.WorkerSupportsCancellation = True Then
            bw.CancelAsync()
        End If
    End Sub

End Class