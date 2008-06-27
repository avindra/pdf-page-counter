Public Class FAQ
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tbBatch As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbPost As System.Windows.Forms.TabPage
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FAQ))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tbBatch = New System.Windows.Forms.TabPage
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbPost = New System.Windows.Forms.TabPage
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.tbBatch.SuspendLayout()
        Me.tbPost.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tbBatch)
        Me.TabControl1.Controls.Add(Me.tbPost)
        Me.TabControl1.Location = New System.Drawing.Point(16, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(600, 512)
        Me.TabControl1.TabIndex = 0
        '
        'tbBatch
        '
        Me.tbBatch.Controls.Add(Me.Label16)
        Me.tbBatch.Controls.Add(Me.Label15)
        Me.tbBatch.Controls.Add(Me.Label8)
        Me.tbBatch.Controls.Add(Me.Label7)
        Me.tbBatch.Controls.Add(Me.Label4)
        Me.tbBatch.Controls.Add(Me.Label2)
        Me.tbBatch.Controls.Add(Me.Label3)
        Me.tbBatch.Controls.Add(Me.Label1)
        Me.tbBatch.Location = New System.Drawing.Point(4, 22)
        Me.tbBatch.Name = "tbBatch"
        Me.tbBatch.Size = New System.Drawing.Size(592, 486)
        Me.tbBatch.TabIndex = 0
        Me.tbBatch.Text = "Building the Batch"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(20, 202)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(552, 48)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Under the ""View"" menu, choose ""Selected Directories"" for a folder report."
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 168)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(560, 48)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "What if I want to see which folders I chose?"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(552, 48)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = resources.GetString("Label4.Text")
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(560, 48)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "What if I want to remove files?"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(552, 56)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "In the ""Mode"" section, select ""Add Folders"", instead of ""One Folder"""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(560, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "What if I want to add another directory?"
        '
        'tbPost
        '
        Me.tbPost.Controls.Add(Me.Label14)
        Me.tbPost.Controls.Add(Me.Label13)
        Me.tbPost.Controls.Add(Me.Label12)
        Me.tbPost.Controls.Add(Me.Label11)
        Me.tbPost.Controls.Add(Me.Label6)
        Me.tbPost.Controls.Add(Me.Label5)
        Me.tbPost.Controls.Add(Me.Label9)
        Me.tbPost.Controls.Add(Me.Label10)
        Me.tbPost.Location = New System.Drawing.Point(4, 22)
        Me.tbPost.Name = "tbPost"
        Me.tbPost.Size = New System.Drawing.Size(592, 486)
        Me.tbPost.TabIndex = 1
        Me.tbPost.Text = "Counting"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(19, 224)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(552, 48)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Its simple as going to ""File"" then choosing ""Save."" Choose where you want to save" & _
            " it, and a spreadsheet will be made accordingly. Alternatively, you can press ""C" & _
            "TRL+S"" on the keyboard."
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(17, 196)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(560, 48)
        Me.Label13.TabIndex = 12
        Me.Label13.Text = "How do I make an Excel spreadsheet?"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(20, 172)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(552, 48)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Under the ""View"" menu, choose ""Report"" for a semi-nice looking report programmed " & _
            "with Crystal Reports."
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(16, 140)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(560, 48)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "How about a nice looking report?"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(20, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(552, 32)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "This is in the ""View"" menu, under statistics."
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(560, 48)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "How do I see the number of files/statistics?"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(20, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(552, 32)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "No it didn't. Naturally, counting thousands of large files would take some time. " & _
            "In the meantime, use the computer minimally, or not at all to avoid a program cr" & _
            "ash."
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(560, 48)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "The program froze!"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Bell MT", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(16, 219)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(560, 48)
        Me.Label15.TabIndex = 6
        Me.Label15.Text = "Can I print?"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(20, 250)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(552, 48)
        Me.Label16.TabIndex = 7
        Me.Label16.Text = "Sure. Its under File > Print. Or you can press ""CTRL + P"""
        '
        'FAQ
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(648, 541)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FAQ"
        Me.Text = "FAQ"
        Me.TabControl1.ResumeLayout(False)
        Me.tbBatch.ResumeLayout(False)
        Me.tbPost.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
