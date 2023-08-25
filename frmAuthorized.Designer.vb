<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAuthorized
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSerialNum = New System.Windows.Forms.TextBox()
        Me.txtAuthorization = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnConfirm = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "您的序號"
        '
        'txtSerialNum
        '
        Me.txtSerialNum.Location = New System.Drawing.Point(116, 28)
        Me.txtSerialNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtSerialNum.Name = "txtSerialNum"
        Me.txtSerialNum.ReadOnly = True
        Me.txtSerialNum.Size = New System.Drawing.Size(303, 30)
        Me.txtSerialNum.TabIndex = 1
        '
        'txtAuthorization
        '
        Me.txtAuthorization.Location = New System.Drawing.Point(116, 79)
        Me.txtAuthorization.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtAuthorization.Name = "txtAuthorization"
        Me.txtAuthorization.Size = New System.Drawing.Size(303, 30)
        Me.txtAuthorization.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 19)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "授權碼"
        '
        'btnConfirm
        '
        Me.btnConfirm.AutoSize = True
        Me.btnConfirm.Location = New System.Drawing.Point(438, 77)
        Me.btnConfirm.Name = "btnConfirm"
        Me.btnConfirm.Size = New System.Drawing.Size(59, 29)
        Me.btnConfirm.TabIndex = 4
        Me.btnConfirm.Text = "確定"
        Me.btnConfirm.UseVisualStyleBackColor = True
        '
        'frmAuthorized
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 133)
        Me.Controls.Add(Me.btnConfirm)
        Me.Controls.Add(Me.txtAuthorization)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSerialNum)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("標楷體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.Name = "frmAuthorized"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "授權檢查"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txtSerialNum As TextBox
    Friend WithEvents txtAuthorization As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnConfirm As Button
End Class
