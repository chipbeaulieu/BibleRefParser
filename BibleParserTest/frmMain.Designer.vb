<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.btnParse = New System.Windows.Forms.Button()
        Me.txtResults = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnUseTopsFile = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pgMain = New System.Windows.Forms.ToolStripProgressBar()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(89, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Bible Reference"
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(218, 73)
        Me.txtReference.Name = "txtReference"
        Me.txtReference.Size = New System.Drawing.Size(421, 26)
        Me.txtReference.TabIndex = 1
        '
        'btnParse
        '
        Me.btnParse.Location = New System.Drawing.Point(645, 71)
        Me.btnParse.Name = "btnParse"
        Me.btnParse.Size = New System.Drawing.Size(98, 37)
        Me.btnParse.TabIndex = 2
        Me.btnParse.Text = "Parse"
        Me.btnParse.UseVisualStyleBackColor = True
        '
        'txtResults
        '
        Me.txtResults.Location = New System.Drawing.Point(93, 134)
        Me.txtResults.Multiline = True
        Me.txtResults.Name = "txtResults"
        Me.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResults.Size = New System.Drawing.Size(647, 369)
        Me.txtResults.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(89, 111)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Results"
        '
        'btnUseTopsFile
        '
        Me.btnUseTopsFile.Location = New System.Drawing.Point(93, 26)
        Me.btnUseTopsFile.Name = "btnUseTopsFile"
        Me.btnUseTopsFile.Size = New System.Drawing.Size(647, 41)
        Me.btnUseTopsFile.TabIndex = 5
        Me.btnUseTopsFile.Text = "Use Tops.Txt"
        Me.btnUseTopsFile.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.pgMain})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 529)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(821, 28)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'pgMain
        '
        Me.pgMain.AutoSize = False
        Me.pgMain.Name = "pgMain"
        Me.pgMain.Size = New System.Drawing.Size(400, 22)
        '
        'frmMain
        '
        Me.AcceptButton = Me.btnParse
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(821, 557)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnUseTopsFile)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtResults)
        Me.Controls.Add(Me.btnParse)
        Me.Controls.Add(Me.txtReference)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmMain"
        Me.Text = "Bible Parser Test"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents btnParse As System.Windows.Forms.Button
    Friend WithEvents txtResults As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnUseTopsFile As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents pgMain As System.Windows.Forms.ToolStripProgressBar

End Class
