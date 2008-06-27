Public Class dirPresent
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

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
    Friend WithEvents lstDir As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstDir = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'lstDir
        '
        Me.lstDir.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstDir.Location = New System.Drawing.Point(0, 0)
        Me.lstDir.Name = "lstDir"
        Me.lstDir.Size = New System.Drawing.Size(568, 524)
        Me.lstDir.TabIndex = 0
        '
        'dirPresent
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(568, 525)
        Me.Controls.Add(Me.lstDir)
        Me.Name = "dirPresent"
        Me.Text = "Directories"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Directories_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        For i = 0 To pdfData.mainDirectories.Length - 1
            lstDir.Items.Add(pdfData.mainDirectories(i))
        Next
    End Sub
End Class
