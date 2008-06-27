Public Class frmStats
    Inherits System.Windows.Forms.Form
    Public mainForm As frmMain
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblLong As System.Windows.Forms.Label
    Friend WithEvents btnDone As System.Windows.Forms.Button
    Friend WithEvents lblFiles As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblLong = New System.Windows.Forms.Label
        Me.lblFiles = New System.Windows.Forms.Label
        Me.btnDone = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Longest File:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Number of Files:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLong
        '
        Me.lblLong.Location = New System.Drawing.Point(168, 16)
        Me.lblLong.Name = "lblLong"
        Me.lblLong.Size = New System.Drawing.Size(344, 24)
        Me.lblLong.TabIndex = 2
        Me.lblLong.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFiles
        '
        Me.lblFiles.Location = New System.Drawing.Point(168, 48)
        Me.lblFiles.Name = "lblFiles"
        Me.lblFiles.Size = New System.Drawing.Size(344, 24)
        Me.lblFiles.TabIndex = 3
        Me.lblFiles.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(227, 86)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(284, 23)
        Me.btnDone.TabIndex = 4
        Me.btnDone.Text = "&Done"
        Me.btnDone.UseVisualStyleBackColor = True
        '
        'frmStats
        '
        Me.AcceptButton = Me.btnDone
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(522, 116)
        Me.Controls.Add(Me.btnDone)
        Me.Controls.Add(Me.lblFiles)
        Me.Controls.Add(Me.lblLong)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(528, 104)
        Me.Name = "frmStats"
        Me.Text = "PDF Statistics"
        Me.ResumeLayout(False)

    End Sub
#End Region
    Private Sub frmStats_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If pdfData.main.Rows.Count = 0 Then
            lblFiles.Text = "You haven't selected any files!"
            lblLong.Text = "You haven't selected any files!"
            Exit Sub
        End If
        Dim i, intLongest As Integer
        Dim excep As Boolean
        Try
            For i = 0 To pdfData.main.Rows.Count - 1
                If pdfData.main.Rows(intLongest).Item(1) < pdfData.main.Rows(i).Item(1) Then
                    intLongest = i
                End If
            Next
        Catch ex As Exception
            excep = True
        End Try
        If excep Then
            lblLong.Text = "You haven't counted the PDF files yet."
        Else
            lblLong.Text = pdfData.main.Rows(intLongest).Item(0) & " (" & pdfData.main.Rows(intLongest).Item(1) & " Pages)"
        End If
        lblFiles.Text = pdfData.main.Rows.Count & " PDF Files"
    End Sub
    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        Me.Close()
    End Sub
End Class
