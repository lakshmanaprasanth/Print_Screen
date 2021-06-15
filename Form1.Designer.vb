<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CreateDocument = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(2, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(127, 83)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "ScreenShot"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CreateDocument
        '
        Me.CreateDocument.Location = New System.Drawing.Point(135, 4)
        Me.CreateDocument.Name = "CreateDocument"
        Me.CreateDocument.Size = New System.Drawing.Size(132, 83)
        Me.CreateDocument.TabIndex = 1
        Me.CreateDocument.Text = "CreateDocument"
        Me.CreateDocument.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(2, 94)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(105, 19)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Add Comment"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(2, 118)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.PlaceholderText = "Comment Box"
        Me.TextBox1.Size = New System.Drawing.Size(264, 23)
        Me.TextBox1.TabIndex = 3
        Me.TextBox1.UseWaitCursor = True
        '
        'Form1
        '
        Me.AccessibleRole = System.Windows.Forms.AccessibleRole.Application
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(268, 147)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CreateDocument)
        Me.Controls.Add(Me.Button1)
        Me.HelpButton = True
        Me.Name = "Form1"
        Me.Text = "Print Screen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents CreateDocument As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents TextBox1 As TextBox
End Class
