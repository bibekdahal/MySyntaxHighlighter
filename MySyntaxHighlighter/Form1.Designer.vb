<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.SyntaxHighlighter1 = New MySyntaxHighlighter.SyntaxHighlighter(Me.components)
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Button1.Location = New System.Drawing.Point(0, 412)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(707, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Refresh All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SyntaxHighlighter1
        '
        Me.SyntaxHighlighter1.AcceptsTab = True
        Me.SyntaxHighlighter1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SyntaxHighlighter1.Location = New System.Drawing.Point(0, 0)
        Me.SyntaxHighlighter1.Name = "SyntaxHighlighter1"
        Me.SyntaxHighlighter1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth
        Me.SyntaxHighlighter1.Size = New System.Drawing.Size(707, 412)
        Me.SyntaxHighlighter1.TabIndex = 0
        Me.SyntaxHighlighter1.Text = ""
        Me.SyntaxHighlighter1.WordWrap = False
        '
        'Button2
        '
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Button2.Location = New System.Drawing.Point(0, 435)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(707, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Paste"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(707, 458)
        Me.Controls.Add(Me.SyntaxHighlighter1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SyntaxHighlighter1 As MySyntaxHighlighter.SyntaxHighlighter
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button

End Class
