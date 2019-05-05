Imports System.Windows.Forms
Imports System.Data.OleDb

Public Class frmMDI
    Friend WithEvents hpAdvancedCHM As System.Windows.Forms.HelpProvider
    Public iYear As Integer

    'Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click
    '    ' Create a new instance of the child form.
    '    Dim ChildForm As New System.Windows.Forms.Form
    '    ' Make it a child of this MDI form before showing it.
    '    ChildForm.MdiParent = Me

    '    m_ChildFormNumber += 1
    '    ChildForm.Text = "Window " & m_ChildFormNumber
    '    ChildForm.Show()
    'End Sub

    'Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
    '    Dim OpenFileDialog As New OpenFileDialog
    '    OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '    OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    '    If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
    '        Dim FileName As String = OpenFileDialog.FileName
    '        ' TODO: Add code here to open the file.
    '    End If
    'End Sub

    'Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
    '    Dim SaveFileDialog As New SaveFileDialog
    '    SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '    SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

    '    If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
    '        Dim FileName As String = SaveFileDialog.FileName
    '        ' TODO: Add code here to save the current contents of the form to a file.
    '    End If
    'End Sub

    'Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
    '    Global.System.Windows.Forms.Application.Exit()
    'End Sub

    'Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
    '    ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    'End Sub

    'Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
    '    ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    'End Sub

    'Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
    '    'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    'End Sub

    'Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StatusBarToolStripMenuItem.Click
    '    Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    'End Sub

    'Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
    '    Me.LayoutMdi(MdiLayout.Cascade)
    'End Sub

    'Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
    '    Me.LayoutMdi(MdiLayout.TileVertical)
    'End Sub

    'Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
    '    Me.LayoutMdi(MdiLayout.TileHorizontal)
    'End Sub

    'Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
    '    Me.LayoutMdi(MdiLayout.ArrangeIcons)
    'End Sub

    'Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
    '    ' Close all child forms of the parent.
    '    For Each ChildForm As Form In Me.MdiChildren
    '        ChildForm.Close()
    '    Next
    'End Sub

    'Private m_ChildFormNumber As Integer = 0

    Public Sub ViewForm(ByVal senderName As String, Optional ByRef LinkForm As frmList = Nothing)
        Dim dt As DataTable
        Dim frm As frmList
        Dim dRow As DataRow
        Dim i As Integer
        Dim d As Date
        On Error GoTo Err_h

        dt = Populate("select * from M_Tables where Menu='" & senderName & "' Order By NomTable")
        Using reader As New DataTableReader(dt)
            If reader.Read Then
                If IsDBNull(reader("Table")) Then Exit Sub
                If reader("Table") = "" Then Exit Sub
                If Not IsDBNull(reader("FormCaption")) And LinkForm Is Nothing Then
                    For Each frm In Me.MdiChildren
                        If frm.Header = reader("FormCaption") Then
                            frm.Activate()
                            Exit Sub
                        End If
                    Next frm
                End If
                frm = New frmList
                frm.IsFitdgw = False
                Me.AddOwnedForm(frm)
                If Not IsDBNull(reader("NoEdit")) Then frm.SetNoEdit(1, reader("NoEdit"))
                If Not IsDBNull(reader("FormCaption")) Then frm.Header = reader("FormCaption")
                If Not IsDBNull(reader("Menu")) Then frm.MenuInvoke = reader("Menu")
                frm.Text = frm.Header
                'If Me.MdiChildren.Length > 0 Then frm.Text += " (просмотр)"
                If Not (LinkForm Is Nothing) Then
                    frm.CreateLinkDataForForm(LinkForm)
                End If
                frm.CreateLinkData(senderName)
                frm.TableName = reader("Table")
                If frm.TableName = "" Then
                    frm = Nothing
                    Exit Sub
                End If
                If Not (LinkForm Is Nothing) Then
                    'If LinkForm.dgwList.SelectedCells.Count = 0 Then Exit Sub
                    'If LinkForm.dgwList2.SelectedCells.Count = 0 And LinkForm.dgwList2.Visible Then Exit Sub
                    PrepareLinkForm(LinkForm, frm)
                End If
                If Not IsDBNull(reader("LinkField")) Then frm.LinkField = reader("LinkField")
                If Not IsDBNull(reader("IsAllowUserToAddRows")) Then
                    frm.dgwList.AllowUserToAddRows = reader("IsAllowUserToAddRows")
                End If
                If Not IsDBNull(reader("mnuAdd")) Then frm.mnuAdd.Visible = reader("mnuAdd")
                If reader.Read Then
                    If Not IsDBNull(reader("IsAllowUserToAddRows")) Then
                        frm.dgwList2.AllowUserToAddRows = reader("IsAllowUserToAddRows")
                    End If
                    If Not IsDBNull(reader("mnuAdd")) Then frm.mnuAdd2.Visible = reader("mnuAdd")
                    If Not IsDBNull(reader("NoEdit")) Then frm.SetNoEdit(2, reader("NoEdit"))
                    If Not IsDBNull(reader("IsDependOnYear")) Then
                        If reader("IsDependOnYear") Then
                            frm.Filter += " AND ID_Year=" + iYear.ToString
                            frm.Text += " (" + iYear.ToString + ")"
                            frm.Header += " (" + iYear.ToString + ")"
                        End If
                    End If
                    If Not IsDBNull(reader("Table")) Then
                        'frmList.dgwList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        'frmList.dgwList.ColumnHeadersHeight = 45
                        frm.TableName2 = reader("Table")
                    End If
                Else
                    'frm.dgwList.AllowUserToAddRows = True
                    frm.TableName2 = ""
                    'frm.mnuAdd.Visible = False
                End If
            Else
                MsgBox("Не найдены таблица для меню '" & senderName & "'")
                Exit Sub
            End If
            frm.MdiParent = Me
            'Me.mnuAdmin.Visible = False
            frm.Size = frm.MdiParent.Size
            'frm.FitdgwHeight()
            frm.GetTablesAttr()
            If frm.IsTable(1, "NoDockFill") Then
                'If mTable(1) = "Q_Uslugi" Or mTable(1) = "Q_DMS" Then
                frm.dgwList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                frm.Panel1.Dock = DockStyle.Left
                'Split1.Width = Me.Width * 2
            End If

            frm.GetColumnWidthFromSettings(1)
            If frm.TableName2 <> "" Then frm.GetColumnWidthFromSettings(2)
            frm.Show()
            frm.IsFitdgw = True
            frm.Fitdgw()
            DeselectAllCells(frm)
        End Using
        Exit Sub
Err_h:
        ErrMess(Err.Description, "frmMDI.ViewForm")
    End Sub


    Private Sub mnuPodr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPodr.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub frmMDI_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        'If Conn.State = ConnectionState.Open Then Conn.Close()
        SaveSetting("MedService", "Settings", "PathDB", PathDB)
        SaveSetting("MedService", "Settings", "Year", iYear)
        MakeCopyDB()
        If Not (modReport.objExcel Is Nothing) Then
            If Not objExcel.Visible Then
                objExcel.Quit()
                objExcel = Nothing
            End If
        End If
    End Sub

    Private Sub frmMDI_HelpRequested(ByVal sender As Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles Me.HelpRequested
        Help.ShowHelp(Me, hpAdvancedCHM.HelpNamespace)
    End Sub

    Private Sub frmMDI_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sYear As String

        PathDB = GetSetting(AppName, "Settings", "PathDB")
        sYear = GetSetting(AppName, "Settings", "Year")
        If sYear = "" Or sYear = "2008" Then
            iYear = 2008
            mnu2008.Checked = True
            mnu2009.Checked = False
        Else
            iYear = CInt(sYear)
            mnu2008.Checked = False
            mnu2009.Checked = True
        End If
        Me.Text = "Учет МедУслуг 1.0" '+ iYear.ToString + " год"
        If PathDB = "" Then
            MsgBox("Не указан путь к базе данных.")
        Else
            Connection = New OleDbConnection(ConnectionString)
            'ADOConn = New ADODB.Connection
            'OpenADOConnection()
        End If
    End Sub

    'Private Sub mnuSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAdmin.Click
    '    frmDB.Show()
    'End Sub

    Private Sub mnuStrahComp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStrahComp.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuVidMedStrah_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuVidMedStrah.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuMedUslugi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMedUslugi.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuStrahPolic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStrahPolic.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuStatCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStatCalc.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuEdIzm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdIzm.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuVidRascheta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuVidRascheta.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFilter.Click
        'frmFilterCalc.Show()
        Dim frm As New frmFilterData
        Dim frmL As frmList
        frmL = Me.ActiveMdiChild
        If frmL.TableName Like "Q_SF*" Then
            frm.Source = "СФ"
            'frm.lblCaption.Text = "Интервал дат"
            'frm.Text = "Выберите интервал дат последней оплаты"
            frm.OK_Button.Text = "OK"
            If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                frm.SetFilter(frmL)
                frmL.ApplyFilter()
            End If
        End If
    End Sub

    'Private Sub mnuServiceAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ViewForm(sender.Name)
    'End Sub
    
    'Private Sub mnuServicesByCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ViewForm(sender.Name)
    'End Sub

    Private Sub mnuRepReestr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRepReestr.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustomers.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuIncomesByPodr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIncomesByPodr.Click
        CreateReport(sender)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub mnuFileDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileDB.Click
        frmDB.Show()
    End Sub

    Private Sub mnuLinkWithAnother_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLinkWithAnother.Click
        Dim frm As frmList

        frm = Me.ActiveMdiChild
        frm.IsLinkingWithAnother = True
        MsgBox("Выберите новую строку в верхней секции для привязки.")
        'frm.LinkWithAnother()
    End Sub

    Private Sub ContentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem.Click
        Help.ShowHelp(Me, hpAdvancedCHM.HelpNamespace)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        'UpdatePodr("R_SoldSt")
        'UpdatePodr("R_SoldAmb")
        'UpdatePodr("R_DMS")
        'UpdatePodr("R_SpecSF")
        frmAbout.ShowDialog()
    End Sub

    Private Sub butSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSave.Click
        Dim frm As frmList

        frm = Me.ActiveMdiChild
        frm.SaveData()
    End Sub

    Private Sub butRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As frmList

        If MsgBox("Восстановить информацию из базы данных? Вся не сохраненная информация будет потеряна.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            frm = Me.ActiveMdiChild
            'frm.RefreshData()
        End If
    End Sub

    Private Sub mnuPreiskurantBas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreiskurantBas.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuPreiskurantOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreiskurantOther.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuSoldAmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSoldAmb.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuSoldSt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSoldSt.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuDMS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDMS.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuPeriods_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPeriods.Click
        ViewForm(sender.Name)
    End Sub
    
    'Private Sub mnuSoldAndDMS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ViewForm(sender.Name)
    'End Sub

    Private Sub mnuSFReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSFReport.Click
        modReport.CreateReport(sender)
    End Sub

    Private Sub mnu2008_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu2008.Click
        iYear = 2008
        mnu2008.Checked = True
        mnu2009.Checked = False
        Me.Text = "Учет МедУслуг 1.0"
    End Sub

    Private Sub mnu2009_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu2009.Click
        iYear = 2009
        mnu2009.Checked = True
        mnu2008.Checked = False
        Me.Text = "Учет МедУслуг 1.0"
    End Sub

    Private Sub mnuNal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNal.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuBeznalFact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuBeznalFact.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuDebug_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDebug.Click
        mnuDebug.Checked = Not mnuDebug.Checked
    End Sub

    Private Sub mnuStom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStom.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuAkt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAkt.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuZakaz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZakaz.Click
        If Not (frmMDI.ActiveForm Is Nothing) Then
            CreateReport(sender)
        Else
            MsgBox("Необходимо открыть список счетов-фактур и выбрать документ.")
        End If
    End Sub

    Private Sub mnuBeznalProekt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBeznalProekt.Click
        ViewForm(sender.Name)
    End Sub

    Private Sub mnuReestrUslugByStat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReestrUslugByStat.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuIncomesByStat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIncomesByStat.Click
        CreateReport(sender)
    End Sub

    Private Sub mnuImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImport.Click
        Dim frm As New frmImport
        frm.Show()
    End Sub

    Private Sub mnuCreateDW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreateDW.Click
        CreateDW(sender)
    End Sub
End Class
