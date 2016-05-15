<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelpLocation
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
        Me.TextBox_OfficeHelp = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_ExcelHelp = New System.Windows.Forms.TextBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(179, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Office VBA 帮助文档所在文件夹"
        '
        'TextBox_OfficeHelp
        '
        Me.TextBox_OfficeHelp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox_OfficeHelp.Location = New System.Drawing.Point(24, 29)
        Me.TextBox_OfficeHelp.Name = "TextBox_OfficeHelp"
        Me.TextBox_OfficeHelp.Size = New System.Drawing.Size(524, 21)
        Me.TextBox_OfficeHelp.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Excel VBA 帮助文档所在文件"
        '
        'TextBox_ExcelHelp
        '
        Me.TextBox_ExcelHelp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox_ExcelHelp.Location = New System.Drawing.Point(24, 73)
        Me.TextBox_ExcelHelp.Name = "TextBox_ExcelHelp"
        Me.TextBox_ExcelHelp.Size = New System.Drawing.Size(524, 21)
        Me.TextBox_ExcelHelp.TabIndex = 1
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(473, 105)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 2
        Me.btnOk.Text = "确定"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'HelpLocation
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 140)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.TextBox_ExcelHelp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_OfficeHelp)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "HelpLocation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HelpLocation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox_OfficeHelp As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox_ExcelHelp As TextBox
    Friend WithEvents btnOk As Button
End Class
