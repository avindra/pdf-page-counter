Imports System.IO
Public Class ReportViewer
    Inherits System.Windows.Forms.Form
    Dim passPrice As String
    Dim passDPI As String
    Dim getCount As Integer
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal DPI_Value As String, ByVal thePrice As String, ByVal PageCount As Integer)
        MyBase.New()
        passPrice = thePrice
        passDPI = DPI_Value
        getCount = PageCount
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
    Friend WithEvents mnuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuNav As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPrev As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNext As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuDone As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents rptView As CrystalDecisions.Windows.Forms.CrystalReportViewer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.rptView = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.mnuMain = New System.Windows.Forms.MenuStrip
        Me.mnuNav = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPrev = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNext = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.mnuDone = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'rptView
        '
        Me.rptView.ActiveViewIndex = -1
        Me.rptView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rptView.DisplayGroupTree = False
        Me.rptView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rptView.Location = New System.Drawing.Point(0, 0)
        Me.rptView.Name = "rptView"
        Me.rptView.SelectionFormula = ""
        Me.rptView.ShowCloseButton = False
        Me.rptView.ShowGroupTreeButton = False
        Me.rptView.ShowRefreshButton = False
        Me.rptView.Size = New System.Drawing.Size(632, 541)
        Me.rptView.TabIndex = 3
        Me.rptView.ViewTimeSelectionFormula = ""
        '
        'mnuMain
        '
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNav})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(632, 24)
        Me.mnuMain.TabIndex = 4
        Me.mnuMain.Text = "MenuStrip1"
        Me.mnuMain.Visible = False
        '
        'mnuNav
        '
        Me.mnuNav.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrev, Me.mnuNext, Me.ToolStripSeparator1, Me.mnuDone})
        Me.mnuNav.Name = "mnuNav"
        Me.mnuNav.Size = New System.Drawing.Size(70, 20)
        Me.mnuNav.Text = "Navigation"
        '
        'mnuPrev
        '
        Me.mnuPrev.Name = "mnuPrev"
        Me.mnuPrev.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.Left), System.Windows.Forms.Keys)
        Me.mnuPrev.Size = New System.Drawing.Size(186, 22)
        Me.mnuPrev.Text = "&Previous Page"
        '
        'mnuNext
        '
        Me.mnuNext.Name = "mnuNext"
        Me.mnuNext.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.Right), System.Windows.Forms.Keys)
        Me.mnuNext.Size = New System.Drawing.Size(186, 22)
        Me.mnuNext.Text = "N&ext Page"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(183, 6)
        '
        'mnuDone
        '
        Me.mnuDone.Name = "mnuDone"
        Me.mnuDone.Size = New System.Drawing.Size(186, 22)
        Me.mnuDone.Text = "&Done"
        '
        'ReportViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 541)
        Me.Controls.Add(Me.rptView)
        Me.Controls.Add(Me.mnuMain)
        Me.MainMenuStrip = Me.mnuMain
        Me.Name = "ReportViewer"
        Me.Text = "PDF Batch Report"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Sub ReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim finalrep As rptMain = New rptMain
        finalrep.SetDataSource(pdfData.main)
        finalrep.SetParameterValue("CostTotal", passPrice)
        finalrep.SetParameterValue("DPIValue", passDPI)
        finalrep.SetParameterValue("FileCount", getCount)
        finalrep.SetParameterValue("Customer", InputBox("Please enter the customer's name:", "Customer Information"))
        finalrep.SetParameterValue("CustomerID", InputBox("Please enter the customer's ID:", "Customer ID", "CC-"))
        rptView.ReportSource = finalrep
    End Sub
    Private Sub previouspage(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrev.Click
        rptView.ShowPreviousPage()
    End Sub
    Private Sub nextpage(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNext.Click
        rptView.ShowNextPage()
    End Sub
    Private Sub done(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDone.Click
        Me.Close()
    End Sub
End Class