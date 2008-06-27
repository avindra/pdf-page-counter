Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmMain
    Inherits System.Windows.Forms.Form
#Region "    Form GUI Code       "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuReset As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents btnPrice As System.Windows.Forms.Button
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuPageCount As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrice As System.Windows.Forms.MenuItem
    Friend WithEvents dlgFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStats As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents chkColor As System.Windows.Forms.CheckBox
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDir As System.Windows.Forms.MenuItem
    Friend WithEvents grpMode As System.Windows.Forms.GroupBox
    Friend WithEvents radAdd As System.Windows.Forms.RadioButton
    Friend WithEvents radNew As System.Windows.Forms.RadioButton
    Friend WithEvents mnuFAQ As System.Windows.Forms.MenuItem
    Friend WithEvents radFile As System.Windows.Forms.RadioButton
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cboDPI As System.Windows.Forms.ComboBox
    Friend WithEvents tipMode As System.Windows.Forms.ToolTip
    Friend WithEvents gridData As System.Windows.Forms.DataGridView
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.MenuItem
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents dlgPrint As System.Windows.Forms.PrintDialog
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents radSND As System.Windows.Forms.RadioButton
    Friend WithEvents grpPrice As System.Windows.Forms.GroupBox
    Friend WithEvents lblPages As System.Windows.Forms.Label
    Friend WithEvents btnCount As System.Windows.Forms.Button
    Friend WithEvents btnFiles As System.Windows.Forms.Button
    Friend WithEvents btnAFile As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSF As System.Windows.Forms.Button
    Friend WithEvents tipPrice As System.Windows.Forms.ToolTip
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuReset = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.mnuSave = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuRpt = New System.Windows.Forms.MenuItem
        Me.mnuStats = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.mnuDir = New System.Windows.Forms.MenuItem
        Me.mnuEdit = New System.Windows.Forms.MenuItem
        Me.mnuPageCount = New System.Windows.Forms.MenuItem
        Me.mnuPrice = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuFAQ = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuAbout = New System.Windows.Forms.MenuItem
        Me.btnPrice = New System.Windows.Forms.Button
        Me.lblCost = New System.Windows.Forms.Label
        Me.dlgFolder = New System.Windows.Forms.FolderBrowserDialog
        Me.chkColor = New System.Windows.Forms.CheckBox
        Me.grpMode = New System.Windows.Forms.GroupBox
        Me.radSND = New System.Windows.Forms.RadioButton
        Me.radFile = New System.Windows.Forms.RadioButton
        Me.radAdd = New System.Windows.Forms.RadioButton
        Me.radNew = New System.Windows.Forms.RadioButton
        Me.cboDPI = New System.Windows.Forms.ComboBox
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.tipMode = New System.Windows.Forms.ToolTip(Me.components)
        Me.gridData = New System.Windows.Forms.DataGridView
        Me.dlgSave = New System.Windows.Forms.SaveFileDialog
        Me.dlgPrint = New System.Windows.Forms.PrintDialog
        Me.grpPrice = New System.Windows.Forms.GroupBox
        Me.lblPages = New System.Windows.Forms.Label
        Me.btnCount = New System.Windows.Forms.Button
        Me.btnFiles = New System.Windows.Forms.Button
        Me.btnAFile = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSF = New System.Windows.Forms.Button
        Me.tipPrice = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpMode.SuspendLayout()
        CType(Me.gridData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPrice.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuView, Me.mnuEdit, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuReset, Me.MenuItem5, Me.mnuSave, Me.MenuItem3, Me.MenuItem1, Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuReset
        '
        Me.mnuReset.Index = 0
        Me.mnuReset.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftR
        Me.mnuReset.Text = "&Reset"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "-"
        '
        'mnuSave
        '
        Me.mnuSave.Index = 2
        Me.mnuSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSave.Text = "&Save "
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.MenuItem3.Text = "&Print"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 4
        Me.MenuItem1.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 5
        Me.mnuExit.Text = "E&xit"
        '
        'mnuView
        '
        Me.mnuView.Index = 1
        Me.mnuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuRpt, Me.mnuStats, Me.MenuItem4, Me.mnuDir})
        Me.mnuView.Text = "&View"
        '
        'mnuRpt
        '
        Me.mnuRpt.Index = 0
        Me.mnuRpt.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuRpt.Text = "R&eport"
        '
        'mnuStats
        '
        Me.mnuStats.Index = 1
        Me.mnuStats.Text = "S&tatistics"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "-"
        '
        'mnuDir
        '
        Me.mnuDir.Index = 3
        Me.mnuDir.Text = "Selected Directories"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 2
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPageCount, Me.mnuPrice})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuPageCount
        '
        Me.mnuPageCount.Index = 0
        Me.mnuPageCount.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftC
        Me.mnuPageCount.Text = "&Copy Page Count"
        '
        'mnuPrice
        '
        Me.mnuPrice.Index = 1
        Me.mnuPrice.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuPrice.Text = "Copy Total &Price"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 3
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFAQ, Me.MenuItem2, Me.mnuAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuFAQ
        '
        Me.mnuFAQ.Index = 0
        Me.mnuFAQ.Text = "FAQ"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 2
        Me.mnuAbout.Text = "&About"
        '
        'btnPrice
        '
        Me.btnPrice.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btnPrice.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrice.Location = New System.Drawing.Point(5, 65)
        Me.btnPrice.Name = "btnPrice"
        Me.btnPrice.Size = New System.Drawing.Size(96, 20)
        Me.btnPrice.TabIndex = 6
        Me.btnPrice.Text = "Pri&ce"
        Me.btnPrice.UseVisualStyleBackColor = False
        '
        'lblCost
        '
        Me.lblCost.Location = New System.Drawing.Point(368, 120)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(128, 24)
        Me.lblCost.TabIndex = 7
        Me.lblCost.Text = "Total Cost"
        Me.lblCost.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.tipPrice.SetToolTip(Me.lblCost, resources.GetString("lblCost.ToolTip"))
        '
        'dlgFolder
        '
        Me.dlgFolder.Description = "Please select a folder containing your PDF files:"
        Me.dlgFolder.ShowNewFolderButton = False
        '
        'chkColor
        '
        Me.chkColor.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkColor.Location = New System.Drawing.Point(5, 43)
        Me.chkColor.Name = "chkColor"
        Me.chkColor.Size = New System.Drawing.Size(96, 16)
        Me.chkColor.TabIndex = 14
        Me.chkColor.Text = "Color?"
        Me.tipPrice.SetToolTip(Me.chkColor, "Select this option if the document batch is in color.")
        '
        'grpMode
        '
        Me.grpMode.Controls.Add(Me.radSND)
        Me.grpMode.Controls.Add(Me.radFile)
        Me.grpMode.Controls.Add(Me.radAdd)
        Me.grpMode.Controls.Add(Me.radNew)
        Me.grpMode.Location = New System.Drawing.Point(13, 8)
        Me.grpMode.Name = "grpMode"
        Me.grpMode.Size = New System.Drawing.Size(366, 40)
        Me.grpMode.TabIndex = 15
        Me.grpMode.TabStop = False
        Me.grpMode.Text = "Mode"
        Me.tipMode.SetToolTip(Me.grpMode, "The File Mode determines the means by which you are building the batch:")
        '
        'radSND
        '
        Me.radSND.Location = New System.Drawing.Point(239, 15)
        Me.radSND.Name = "radSND"
        Me.radSND.Size = New System.Drawing.Size(92, 18)
        Me.radSND.TabIndex = 4
        Me.radSND.TabStop = True
        Me.radSND.Text = "Sub-Folders!"
        Me.tipMode.SetToolTip(Me.radSND, "Include Sub-Folders: This seek-and-destroy method will find EVERY PDF file," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "even" & _
                " in the folders within folders for counting.")
        Me.radSND.UseVisualStyleBackColor = True
        '
        'radFile
        '
        Me.radFile.AutoSize = True
        Me.radFile.Location = New System.Drawing.Point(172, 15)
        Me.radFile.Name = "radFile"
        Me.radFile.Size = New System.Drawing.Size(68, 17)
        Me.radFile.TabIndex = 3
        Me.radFile.TabStop = True
        Me.radFile.Text = "Add Files"
        Me.tipMode.SetToolTip(Me.radFile, "Add Files: This will append your batch with single files that you select.")
        Me.radFile.UseVisualStyleBackColor = True
        '
        'radAdd
        '
        Me.radAdd.Location = New System.Drawing.Point(93, 16)
        Me.radAdd.Name = "radAdd"
        Me.radAdd.Size = New System.Drawing.Size(83, 16)
        Me.radAdd.TabIndex = 2
        Me.radAdd.Text = "A&dd Folders"
        Me.tipMode.SetToolTip(Me.radAdd, "Add Folders: This will append your batch with a folder that you select.")
        '
        'radNew
        '
        Me.radNew.Checked = True
        Me.radNew.Location = New System.Drawing.Point(16, 16)
        Me.radNew.Name = "radNew"
        Me.radNew.Size = New System.Drawing.Size(82, 16)
        Me.radNew.TabIndex = 1
        Me.radNew.TabStop = True
        Me.radNew.Text = "&One Folder"
        Me.tipMode.SetToolTip(Me.radNew, "One Folder: This will clear the entire batch and replace it with PDF files in the" & _
                " one folder you choose.")
        '
        'cboDPI
        '
        Me.cboDPI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDPI.FormattingEnabled = True
        Me.cboDPI.Items.AddRange(New Object() {"100 DPI", "150 DPI", "200 DPI", "240 DPI", "300 DPI", "400 DPI", "600 DPI"})
        Me.cboDPI.Location = New System.Drawing.Point(5, 16)
        Me.cboDPI.Name = "cboDPI"
        Me.cboDPI.Size = New System.Drawing.Size(96, 21)
        Me.cboDPI.TabIndex = 18
        Me.tipPrice.SetToolTip(Me.cboDPI, resources.GetString("cboDPI.ToolTip"))
        '
        'dlgOpen
        '
        Me.dlgOpen.Filter = "PDF Files|*.pdf|All files|*.*"
        Me.dlgOpen.Multiselect = True
        Me.dlgOpen.Title = "Select Single PDF Files to Add to the Batch"
        '
        'tipMode
        '
        Me.tipMode.ToolTipTitle = "Tell me about the file modes..."
        '
        'gridData
        '
        Me.gridData.AllowUserToAddRows = False
        Me.gridData.AllowUserToResizeColumns = False
        Me.gridData.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Yellow
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.gridData.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.gridData.BackgroundColor = System.Drawing.Color.White
        Me.gridData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridData.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Red
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridData.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.gridData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridData.Cursor = System.Windows.Forms.Cursors.Cross
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.DarkGoldenrod
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.gridData.DefaultCellStyle = DataGridViewCellStyle3
        Me.gridData.Location = New System.Drawing.Point(14, 146)
        Me.gridData.Name = "gridData"
        Me.gridData.ReadOnly = True
        Me.gridData.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.gridData.RowHeadersWidth = 20
        Me.gridData.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.Yellow
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black
        Me.gridData.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.gridData.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Yellow
        Me.gridData.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black
        Me.gridData.RowTemplate.Height = 20
        Me.gridData.Size = New System.Drawing.Size(482, 336)
        Me.gridData.TabIndex = 19
        '
        'dlgSave
        '
        Me.dlgSave.Filter = "Excel 2003 Spreadsheet|*.xls|Excel 2007 Spreadsheet|*.xlsx|CSV File|*.csv"
        '
        'dlgPrint
        '
        Me.dlgPrint.UseEXDialog = True
        '
        'grpPrice
        '
        Me.grpPrice.Controls.Add(Me.cboDPI)
        Me.grpPrice.Controls.Add(Me.chkColor)
        Me.grpPrice.Controls.Add(Me.btnPrice)
        Me.grpPrice.Location = New System.Drawing.Point(389, 3)
        Me.grpPrice.Name = "grpPrice"
        Me.grpPrice.Size = New System.Drawing.Size(109, 100)
        Me.grpPrice.TabIndex = 21
        Me.grpPrice.TabStop = False
        Me.grpPrice.Text = "Pricing"
        '
        'lblPages
        '
        Me.lblPages.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPages.Location = New System.Drawing.Point(16, 58)
        Me.lblPages.Name = "lblPages"
        Me.lblPages.Size = New System.Drawing.Size(228, 48)
        Me.lblPages.TabIndex = 2
        Me.lblPages.Text = "0 Pages"
        Me.lblPages.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCount
        '
        Me.btnCount.BackColor = System.Drawing.Color.Goldenrod
        Me.btnCount.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCount.Location = New System.Drawing.Point(250, 58)
        Me.btnCount.Name = "btnCount"
        Me.btnCount.Size = New System.Drawing.Size(129, 48)
        Me.btnCount.TabIndex = 4
        Me.btnCount.Text = "&Count"
        Me.btnCount.UseVisualStyleBackColor = False
        '
        'btnFiles
        '
        Me.btnFiles.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btnFiles.Location = New System.Drawing.Point(16, 112)
        Me.btnFiles.Name = "btnFiles"
        Me.btnFiles.Size = New System.Drawing.Size(88, 24)
        Me.btnFiles.TabIndex = 0
        Me.btnFiles.Text = "Select F&older"
        Me.btnFiles.UseVisualStyleBackColor = False
        '
        'btnAFile
        '
        Me.btnAFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAFile.ForeColor = System.Drawing.Color.Black
        Me.btnAFile.Location = New System.Drawing.Point(16, 112)
        Me.btnAFile.Name = "btnAFile"
        Me.btnAFile.Size = New System.Drawing.Size(88, 24)
        Me.btnAFile.TabIndex = 17
        Me.btnAFile.Text = "Add Files"
        Me.btnAFile.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.Red
        Me.btnAdd.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAdd.Location = New System.Drawing.Point(16, 112)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(88, 24)
        Me.btnAdd.TabIndex = 16
        Me.btnAdd.Text = "Add Folder&s"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'btnSF
        '
        Me.btnSF.BackColor = System.Drawing.Color.Yellow
        Me.btnSF.Location = New System.Drawing.Point(16, 112)
        Me.btnSF.Name = "btnSF"
        Me.btnSF.Size = New System.Drawing.Size(88, 24)
        Me.btnSF.TabIndex = 20
        Me.btnSF.Text = "Sub-Folders"
        Me.btnSF.UseVisualStyleBackColor = False
        '
        'tipPrice
        '
        Me.tipPrice.ToolTipTitle = "Pricing Help"
        '
        'frmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 473)
        Me.Controls.Add(Me.grpMode)
        Me.Controls.Add(Me.grpPrice)
        Me.Controls.Add(Me.btnSF)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnAFile)
        Me.Controls.Add(Me.gridData)
        Me.Controls.Add(Me.btnFiles)
        Me.Controls.Add(Me.lblCost)
        Me.Controls.Add(Me.btnCount)
        Me.Controls.Add(Me.lblPages)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.mnuMain
        Me.MinimumSize = New System.Drawing.Size(512, 520)
        Me.Name = "frmMain"
        Me.Text = "PDF Page Counter"
        Me.TransparencyKey = System.Drawing.Color.Magenta
        Me.grpMode.ResumeLayout(False)
        Me.grpMode.PerformLayout()
        CType(Me.gridData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPrice.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "    Add Files           "
    Private Sub btnFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFiles.Click
        Dim response As DialogResult = dlgFolder.ShowDialog()
        If response = DialogResult.Cancel Then Exit Sub

        Dim storefile As Directory
        Dim files As String()

        files = storefile.GetFiles(dlgFolder.SelectedPath, "*.pdf")

        If Not files.Length > 0 Then
            MessageBox.Show("Folder has no PDF Files in it.")
            Exit Sub
        End If

        ReDim pdfData.mainDirectories(0)
        pdfData.mainDirectories(0) = dlgFolder.SelectedPath()

        lblPages.Text = "0 Pages"
        lblCost.Text = "Total Cost"
        pdfData.main.Rows.Clear()
        Dim i As Integer
        For i = 0 To files.Length - 1
            pdfData.main.Rows.Add(pdfData.main.NewRow())
            pdfData.main.Rows(i).Item(0) = files(i)
        Next
        Me.AcceptButton = btnCount
        dlgFolder.SelectedPath = Nothing
    End Sub
    Private Sub SeekAndDestroy(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF.Click
        Dim response As DialogResult = dlgFolder.ShowDialog()
        If response = DialogResult.Cancel Then Exit Sub
        Dim storefile As New DirectoryInfo(dlgFolder.SelectedPath)
        Dim files As String() = {"null"}

        ListFiles(storefile, files)

        If Not files.Length > 0 Then
            MessageBox.Show("Folder and all sub-folders have no PDF Files in them.")
            Exit Sub
        End If

        ReDim pdfData.mainDirectories(0)
        pdfData.mainDirectories(0) = dlgFolder.SelectedPath()

        lblPages.Text = "0 Pages"
        lblCost.Text = "Total Cost"
        pdfData.main.Rows.Clear()
        Dim i As Integer
        For i = 0 To files.Length - 1
            pdfData.main.Rows.Add(pdfData.main.NewRow())
            pdfData.main.Rows(i).Item(0) = files(i)
        Next
        pdfData.main.Rows.RemoveAt(pdfData.main.Rows.Count - 1)
        Me.AcceptButton = btnCount
        dlgFolder.SelectedPath = Nothing
    End Sub
    Private Sub ListFiles(ByVal dir_info As DirectoryInfo, ByRef thefiles() As String)
        ' Get the files in this directory.
        Dim fs_infos() As FileInfo = dir_info.GetFiles("*.pdf")
        For Each fs_info As FileInfo In fs_infos
            ReDim Preserve thefiles(thefiles.Length)
            thefiles(thefiles.Length - 2) = fs_info.FullName
        Next fs_info
        fs_infos = Nothing

        ' Search subdirectories.
        Dim subdirs() As DirectoryInfo = dir_info.GetDirectories()
        For Each subdir As DirectoryInfo In subdirs
            ListFiles(subdir, thefiles)
        Next subdir
        ReDim Preserve thefiles(thefiles.Length - 1)
    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim response As DialogResult = dlgFolder.ShowDialog()
        If response = DialogResult.Cancel Then Exit Sub
        Dim i, b As Short
        For Each DirPath As String In pdfData.mainDirectories
            If dlgFolder.SelectedPath = DirPath Then
                MessageBox.Show("You have already selected this directory for counting.", "Duplicate Detected")
                Exit Sub
            End If
        Next
        Dim storefile As Directory
        Dim files As String()

        files = storefile.GetFiles(dlgFolder.SelectedPath, "*.pdf")

        If Not files.Length > 0 Then
            MessageBox.Show("Folder has no PDF Files in it.")
            Exit Sub
        End If

        ReDim Preserve pdfData.mainDirectories(pdfData.mainDirectories.Length)
        pdfData.mainDirectories(pdfData.mainDirectories.Length - 1) = dlgFolder.SelectedPath

        Dim prev As Integer = pdfData.main.Rows.Count
        For Each File As String In files
            pdfData.main.Rows.Add(pdfData.main.NewRow())
        Next
        For i = prev To pdfData.main.Rows.Count - 1
            pdfData.main.Rows(i).Item(0) = files(b)
            b = b + 1
        Next
        dlgFolder.SelectedPath = Nothing
        Me.AcceptButton = btnCount
    End Sub
    Private Sub btnAFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAFile.Click
        Dim response As DialogResult = dlgOpen.ShowDialog()
        If response = DialogResult.Cancel Then Exit Sub
        Dim i, b As Short

        Dim prev As Integer = pdfData.main.Rows.Count
        For Each File As String In dlgOpen.FileNames
            pdfData.main.Rows.Add(pdfData.main.NewRow())
        Next
        For i = prev To pdfData.main.Rows.Count - 1
            pdfData.main.Rows(i).Item(0) = dlgOpen.FileNames(b)
            b = b + 1
        Next
        Me.AcceptButton = btnCount
    End Sub
#End Region
#Region "    Count Pages       "
    Function CountPages(ByVal FileName As String) As Integer
        Dim file As New FileStream(FileName, FileMode.Open)
        Dim reader As New StreamReader(file)
        Dim Data As String = reader.ReadToEnd
        file.Close()
        reader.Close()
        Dim myMatch As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Data, "/Type\s*/Page[^s]")
        Return myMatch.Count.ToString
    End Function
    Private Sub btnCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCount.Click
        Try
            Dim totalCount As Integer
            For Each Row As DataRow In pdfData.main.Rows
                If Row.Item(1).ToString = Nothing Then
                    Row.Item(1) = CountPages(Row.Item(0))
                    totalCount = Row.Item(1) + totalCount
                Else
                    totalCount = Row.Item(1) + totalCount
                End If
            Next
            lblPages.Text = totalCount & " Pages"
        Catch ex As Exception
            MessageBox.Show("You have not selected any files to count!", "Please read...", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End Try
    End Sub
#End Region
#Region "    Mini Helpers     "
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub
    Private Sub mnuReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReset.Click
        ReDim pdfData.mainDirectories(-1)
        lblPages.Text = "0 Pages"
        lblCost.Text = "Total Cost"
    End Sub
    Private Sub mnuPageCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPageCount.Click
        My.Computer.Clipboard.SetText(lblPages.Text.Remove(lblPages.Text.Length - 6, 6))
    End Sub
    Private Sub mnuPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrice.Click
        My.Computer.Clipboard.SetText(lblCost.Text)
    End Sub
    Private Sub radNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radNew.CheckedChanged, radAdd.CheckedChanged, radFile.CheckedChanged
        If radAdd.Checked Then
            btnAdd.Visible = True
            btnAFile.Visible = False
            btnFiles.Visible = False
            btnSF.Visible = False
            Me.AcceptButton = btnAdd
        ElseIf radNew.Checked Then
            btnAdd.Visible = False
            btnAFile.Visible = False
            btnFiles.Visible = True
            btnSF.Visible = False
            Me.AcceptButton = btnFiles
        ElseIf radFile.Checked Then
            btnAFile.Visible = True
            btnAdd.Visible = False
            btnFiles.Visible = False
            btnSF.Visible = False
            Me.AcceptButton = btnAFile
        ElseIf radSND.Checked Then
            btnSF.Visible = True
            btnFiles.Visible = False
            btnAdd.Visible = False
            btnAFile.Visible = False
            Me.AcceptButton = btnSF
        End If
    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        btnAdd.Visible = False
        btnAFile.Visible = False
        cboDPI.Text = "100 DPI"
        pdfData.Files.DataType = System.Type.GetType("System.String")
        pdfData.main.Columns.Add(Files)
        pdfData.Pages.DataType = System.Type.GetType("System.Int32")
        pdfData.main.Columns.Add(Pages)
        gridData.DataSource() = pdfData.main
        gridData.Columns(0).Width = Me.Width - 120
        gridData.Columns(1).Width = 50
    End Sub
#End Region
#Region "    Pricing   "
    Private Sub btnPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrice.Click
        Dim getCount As Integer = lblPages.Text.Remove(lblPages.Text.Length - 6, 6)
        Dim PageFactor As Integer
        Select Case getCount
            Case 0 To 999
                PageFactor = 0
            Case 1000 To 9999
                PageFactor = 1
            Case 10000 To 99999
                PageFactor = 2
            Case 100000 To 499999
                PageFactor = 3
            Case 500000 To 999999
                PageFactor = 4
            Case Else
                PageFactor = 5
        End Select
        lblCost.Text = Format(getCount * DPIGrab(PageFactor) + getCount * 0.15, "Currency")
        Me.AcceptButton = btnFiles
    End Sub
    Function DPIGrab(ByVal DPI As Integer) As Decimal
        Select Case chkColor.Checked
            Case False
                Select Case cboDPI.Text
                    Case "100 DPI"
                        Dim DPIFactor() As Decimal = {0.055, 0.05, 0.045, 0.04, 0.03, 0.025}
                        Return DPIFactor(DPI)
                    Case "150 DPI"
                        Dim DPIFactor() As Decimal = {0.065, 0.6, 0.055, 0.045, 0.035, 0.03}
                        Return DPIFactor(DPI)
                    Case "200 DPI"
                        Dim DPIFactor() As Decimal = {0.075, 0.7, 0.06, 0.05, 0.045, 0.04}
                        Return DPIFactor(DPI)
                    Case "240 DPI"
                        Dim DPIFactor() As Decimal = {0.08, 0.75, 0.065, 0.055, 0.05, 0.045}
                        Return DPIFactor(DPI)
                    Case "300 DPI"
                        Dim DPIFactor() As Decimal = {0.09, 0.85, 0.075, 0.065, 0.06, 0.055}
                        Return DPIFactor(DPI)
                    Case "400 DPI"
                        Dim DPIFactor() As Decimal = {0.95, 0.9, 0.08, 0.07, 0.65, 0.06}
                        Return DPIFactor(DPI)
                    Case "600 DPI"
                        Dim DPIFactor() As Decimal = {0.105, 0.1, 0.09, 0.08, 0.75, 0.07}
                        Return DPIFactor(DPI)
                End Select
            Case True
                Select Case cboDPI.Text
                    Case "100 DPI"
                        Dim DPIFactor() As Decimal = {0.083, 0.075, 0.068, 0.06, 0.45, 0.038}
                        Return DPIFactor(DPI)
                    Case "150 DPI"
                        Dim DPIFactor() As Decimal = {0.098, 0.09, 0.083, 0.068, 0.053, 0.045}
                        Return DPIFactor(DPI)
                    Case "200 DPI"
                        Dim DPIFactor() As Decimal = {0.113, 0.105, 0.09, 0.075, 0.068, 0.06}
                        Return DPIFactor(DPI)
                    Case "240 DPI"
                        Dim DPIFactor() As Decimal = {0.12, 0.113, 0.098, 0.083, 0.075, 0.068}
                        Return DPIFactor(DPI)
                    Case "300 DPI"
                        Dim DPIFactor() As Decimal = {0.135, 0.128, 0.113, 0.098, 0.09, 0.083}
                        Return DPIFactor(DPI)
                    Case "400 DPI"
                        Dim DPIFactor() As Decimal = {0.143, 0.135, 0.12, 0.105, 0.098, 0.09}
                        Return DPIFactor(DPI)
                    Case "600 DPI"
                        Dim DPIFactor() As Decimal = {0.158, 0.15, 0.135, 0.12, 0.113, 0.105}
                        Return DPIFactor(DPI)
                End Select
        End Select
    End Function
#End Region
#Region "    Window Callers   "
    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        Dim about As New About
        about.ShowDialog()
    End Sub
    Private Sub mnuFAQ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFAQ.Click
        Dim faqWin As New FAQ
        faqWin.Show()
    End Sub
    Private Sub mnuStats_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStats.Click
        Dim stats As New frmStats()
        stats.ShowDialog()
    End Sub
    Private Sub ViewReport(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRpt.Click
        If lblCost.Text = "Total Cost" Then
            MessageBox.Show("You have not calculated the price yet!", "Pay attention!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        ElseIf lblPages.Text = "0 Pages" Then
            MessageBox.Show("You have not calculated the page count yet!", "Pay attention!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim viewer As New ReportViewer(cboDPI.Text, lblCost.Text, lblPages.Text.Remove(lblPages.Text.Length - 6, 6))
        viewer.Show()
    End Sub
    Private Sub ViewDirs(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDir.Click
        Dim showDir As New dirPresent
        showDir.ShowDialog()
    End Sub
#End Region
#Region "    Excel/Printing/Saving        "
    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        dlgSave.FileName = My.Computer.Clock.LocalTime.Month & " " & My.Computer.Clock.LocalTime.Day & ", " & My.Computer.Clock.LocalTime.Year
        Dim reply As DialogResult = dlgSave.ShowDialog()
        If reply = DialogResult.Cancel Then Exit Sub
        Dim oApp As New Excel.Application()
        Dim oBooks As Excel.Workbooks = oApp.Workbooks
        Dim oBook As Excel.Workbook = oBooks.Add
        Dim oSheet As Excel.Worksheet = oApp.ActiveSheet
        oSheet.Cells(1, 1) = "Files"
        oSheet.Cells(1, 2) = "Pages"
        oSheet.Name = "PDF Data"
        oSheet.Columns.Item(1).ColumnWidth = 80
        oSheet.Columns.Item(2).ColumnWidth = 5
        oBook.Sheets(1).Delete()
        oBook.Sheets(2).Delete()
        Dim i As Integer
        For i = 0 To pdfData.main.Rows.Count - 1
            oSheet.Cells(i + 2, 1) = pdfData.main.Rows(i).Item(0)
            oSheet.Cells(i + 2, 2) = pdfData.main.Rows(i).Item(1)
        Next
        NAR(oSheet)
        Select Case dlgSave.FilterIndex()
            Case 1
                oBook.SaveAs(dlgSave.FileName, Excel.XlFileFormat.xlExcel8)
            Case 2
                oBook.SaveAs(dlgSave.FileName)
            Case 3
                oBook.SaveAs(dlgSave.FileName, Excel.XlFileFormat.xlCSV)
        End Select
        oBook.Close(False)
        NAR(oBook)
        NAR(oBooks)
        oApp.Quit()
        NAR(oApp)
    End Sub
    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
    Private Sub print(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim reply As DialogResult = dlgPrint.ShowDialog()
        If reply = DialogResult.Cancel Then Exit Sub
        Dim oApp As New Excel.Application()
        Dim oBooks As Excel.Workbooks = oApp.Workbooks
        Dim oBook As Excel.Workbook = oBooks.Add
        Dim oSheet As Excel.Worksheet = oApp.ActiveSheet
        oSheet.Cells(1, 1) = "Files"
        oSheet.Cells(1, 2) = "Pages"
        oSheet.Name = "PDF Data"
        oSheet.Columns.Item(1).ColumnWidth = 80
        oSheet.Columns.Item(2).ColumnWidth = 5
        oBook.Sheets(1).Delete()
        oBook.Sheets(2).Delete()
        Dim i As Integer
        For i = 0 To pdfData.main.Rows.Count - 1
            oSheet.Cells(i + 2, 1) = pdfData.main.Rows(i).Item(0)
            oSheet.Cells(i + 2, 2) = pdfData.main.Rows(i).Item(1)
        Next
        oBook.PrintOutEx(, , dlgPrint.PrinterSettings.Copies, , dlgPrint.PrinterSettings.PrinterName, , dlgPrint.PrinterSettings.Collate, dlgPrint.PrintToFile)
        NAR(oSheet)
        oBook.Close(False)
        NAR(oBook)
        NAR(oBooks)
        oApp.Quit()
        NAR(oApp)
    End Sub
#End Region
#Region "    Form Metrics      "
    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        gridData.Width = Me.Width - 30
        gridData.Height = Me.Height - 200
        Try
            gridData.Columns(0).Width = Me.Width - 120
        Catch
        End Try
    End Sub
#End Region
End Class