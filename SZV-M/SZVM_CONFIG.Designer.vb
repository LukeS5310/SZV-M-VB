<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SZVM_CONFIG
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
        Me.BTN_CLEAN_ALL = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BTN_CLEAN_ALL
        '
        Me.BTN_CLEAN_ALL.Location = New System.Drawing.Point(13, 13)
        Me.BTN_CLEAN_ALL.Name = "BTN_CLEAN_ALL"
        Me.BTN_CLEAN_ALL.Size = New System.Drawing.Size(159, 23)
        Me.BTN_CLEAN_ALL.TabIndex = 0
        Me.BTN_CLEAN_ALL.Text = "Полная очистка базы"
        Me.BTN_CLEAN_ALL.UseVisualStyleBackColor = True
        '
        'SZVM_CONFIG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(184, 47)
        Me.Controls.Add(Me.BTN_CLEAN_ALL)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "SZVM_CONFIG"
        Me.Text = "Настройки"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BTN_CLEAN_ALL As Button
End Class
