<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer = New System.Windows.Forms.SplitContainer
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cldr1 = New System.Windows.Forms.MonthCalendar
        Me.dgwList = New System.Windows.Forms.DataGridView
        Me.mnuCon = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAdd = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDel = New System.Windows.Forms.ToolStripMenuItem
        Me.dgwSumma = New System.Windows.Forms.DataGridView
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.cldr2 = New System.Windows.Forms.MonthCalendar
        Me.dgwList2 = New System.Windows.Forms.DataGridView
        Me.mnuCon2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAdd2 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDel2 = New System.Windows.Forms.ToolStripMenuItem
        Me.dgwSumma2 = New System.Windows.Forms.DataGridView
        Me.mnuCopy = New System.Windows.Forms.ToolStripMenuItem
        Me.SplitContainer.Panel1.SuspendLayout()
        Me.SplitContainer.Panel2.SuspendLayout()
        Me.SplitContainer.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgwList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.mnuCon.SuspendLayout()
        CType(Me.dgwSumma, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.dgwList2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.mnuCon2.SuspendLayout()
        CType(Me.dgwSumma2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer
        '
        Me.SplitContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer.Name = "SplitContainer"
        Me.SplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer.Panel1
        '
        Me.SplitContainer.Panel1.AutoScroll = True
        Me.SplitContainer.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer.Panel1.Controls.Add(Me.dgwSumma)
        '
        'SplitContainer.Panel2
        '
        Me.SplitContainer.Panel2.AutoScroll = True
        Me.SplitContainer.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer.Panel2.Controls.Add(Me.dgwSumma2)
        Me.SplitContainer.Size = New System.Drawing.Size(695, 348)
        Me.SplitContainer.SplitterDistance = 153
        Me.SplitContainer.SplitterWidth = 3
        Me.SplitContainer.TabIndex = 5
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cldr1)
        Me.Panel1.Controls.Add(Me.dgwList)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(695, 131)
        Me.Panel1.TabIndex = 5
        '
        'cldr1
        '
        Me.cldr1.Location = New System.Drawing.Point(270, 132)
        Me.cldr1.Margin = New System.Windows.Forms.Padding(7, 7, 7, 7)
        Me.cldr1.Name = "cldr1"
        Me.cldr1.TabIndex = 4
        Me.cldr1.Visible = False
        '
        'dgwList
        '
        Me.dgwList.AllowUserToAddRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        DataGridViewCellStyle1.NullValue = Nothing
        Me.dgwList.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgwList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgwList.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgwList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgwList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgwList.ContextMenuStrip = Me.mnuCon
        Me.dgwList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgwList.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgwList.Location = New System.Drawing.Point(0, 0)
        Me.dgwList.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgwList.MultiSelect = False
        Me.dgwList.Name = "dgwList"
        Me.dgwList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgwList.RowTemplate.Height = 24
        Me.dgwList.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgwList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgwList.ShowRowErrors = False
        Me.dgwList.Size = New System.Drawing.Size(695, 131)
        Me.dgwList.TabIndex = 1
        '
        'mnuCon
        '
        Me.mnuCon.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAdd, Me.mnuDel, Me.mnuCopy})
        Me.mnuCon.Name = "mnuCon"
        Me.mnuCon.Size = New System.Drawing.Size(153, 92)
        '
        'mnuAdd
        '
        Me.mnuAdd.Name = "mnuAdd"
        Me.mnuAdd.Size = New System.Drawing.Size(152, 22)
        Me.mnuAdd.Text = "&Добавить"
        '
        'mnuDel
        '
        Me.mnuDel.Name = "mnuDel"
        Me.mnuDel.Size = New System.Drawing.Size(152, 22)
        Me.mnuDel.Text = "&Удалить"
        '
        'dgwSumma
        '
        Me.dgwSumma.AllowUserToAddRows = False
        Me.dgwSumma.AllowUserToDeleteRows = False
        Me.dgwSumma.AllowUserToResizeColumns = False
        Me.dgwSumma.AllowUserToResizeRows = False
        Me.dgwSumma.ColumnHeadersVisible = False
        Me.dgwSumma.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgwSumma.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgwSumma.Enabled = False
        Me.dgwSumma.EnableHeadersVisualStyles = False
        Me.dgwSumma.Location = New System.Drawing.Point(0, 131)
        Me.dgwSumma.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgwSumma.Name = "dgwSumma"
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        Me.dgwSumma.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgwSumma.RowTemplate.Height = 24
        Me.dgwSumma.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.dgwSumma.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgwSumma.Size = New System.Drawing.Size(695, 22)
        Me.dgwSumma.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.cldr2)
        Me.Panel2.Controls.Add(Me.dgwList2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(695, 170)
        Me.Panel2.TabIndex = 0
        '
        'cldr2
        '
        Me.cldr2.Location = New System.Drawing.Point(252, 19)
        Me.cldr2.Margin = New System.Windows.Forms.Padding(7, 7, 7, 7)
        Me.cldr2.Name = "cldr2"
        Me.cldr2.TabIndex = 9
        Me.cldr2.Visible = False
        '
        'dgwList2
        '
        Me.dgwList2.AllowUserToAddRows = False
        Me.dgwList2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgwList2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
        Me.dgwList2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgwList2.ContextMenuStrip = Me.mnuCon2
        Me.dgwList2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgwList2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgwList2.Location = New System.Drawing.Point(0, 0)
        Me.dgwList2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgwList2.MultiSelect = False
        Me.dgwList2.Name = "dgwList2"
        Me.dgwList2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgwList2.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgwList2.RowTemplate.Height = 24
        Me.dgwList2.Size = New System.Drawing.Size(695, 170)
        Me.dgwList2.TabIndex = 8
        Me.dgwList2.Visible = False
        '
        'mnuCon2
        '
        Me.mnuCon2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAdd2, Me.mnuDel2})
        Me.mnuCon2.Name = "mnuCon"
        Me.mnuCon2.Size = New System.Drawing.Size(136, 48)
        '
        'mnuAdd2
        '
        Me.mnuAdd2.Name = "mnuAdd2"
        Me.mnuAdd2.Size = New System.Drawing.Size(135, 22)
        Me.mnuAdd2.Text = "&Добавить"
        Me.mnuAdd2.Visible = False
        '
        'mnuDel2
        '
        Me.mnuDel2.Name = "mnuDel2"
        Me.mnuDel2.Size = New System.Drawing.Size(135, 22)
        Me.mnuDel2.Text = "&Удалить"
        '
        'dgwSumma2
        '
        Me.dgwSumma2.AllowUserToAddRows = False
        Me.dgwSumma2.AllowUserToDeleteRows = False
        Me.dgwSumma2.AllowUserToResizeColumns = False
        Me.dgwSumma2.AllowUserToResizeRows = False
        Me.dgwSumma2.ColumnHeadersVisible = False
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgwSumma2.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgwSumma2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgwSumma2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgwSumma2.Enabled = False
        Me.dgwSumma2.EnableHeadersVisualStyles = False
        Me.dgwSumma2.Location = New System.Drawing.Point(0, 170)
        Me.dgwSumma2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.dgwSumma2.Name = "dgwSumma2"
        Me.dgwSumma2.ReadOnly = True
        Me.dgwSumma2.RowTemplate.Height = 24
        Me.dgwSumma2.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.dgwSumma2.Size = New System.Drawing.Size(695, 22)
        Me.dgwSumma2.TabIndex = 10
        Me.dgwSumma2.Visible = False
        '
        'mnuCopy
        '
        Me.mnuCopy.Name = "mnuCopy"
        Me.mnuCopy.Size = New System.Drawing.Size(152, 22)
        Me.mnuCopy.Text = "Копировать"
        Me.mnuCopy.Visible = False
        '
        'frmList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(695, 348)
        Me.Controls.Add(Me.SplitContainer)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "frmList"
        Me.Text = "frmList"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer.Panel1.ResumeLayout(False)
        Me.SplitContainer.Panel2.ResumeLayout(False)
        Me.SplitContainer.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgwList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.mnuCon.ResumeLayout(False)
        CType(Me.dgwSumma, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgwList2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.mnuCon2.ResumeLayout(False)
        CType(Me.dgwSumma2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer As System.Windows.Forms.SplitContainer
    Friend WithEvents dgwSumma As System.Windows.Forms.DataGridView
    Friend WithEvents cldr1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents mnuCon As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCon2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAdd2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cldr2 As System.Windows.Forms.MonthCalendar
    Friend WithEvents dgwSumma2 As System.Windows.Forms.DataGridView
    Friend WithEvents dgwList2 As System.Windows.Forms.DataGridView
    Friend WithEvents mnuDel2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dgwList As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents mnuCopy As System.Windows.Forms.ToolStripMenuItem
End Class
