<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScaleQuery
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cmbProduct = New System.Windows.Forms.ComboBox()
        Me.cmbCarNo = New System.Windows.Forms.ComboBox()
        Me.cmbManufacturer = New System.Windows.Forms.ComboBox()
        Me.cmbCustomer = New System.Windows.Forms.ComboBox()
        Me.btnQuery = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 27)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "客  戶"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 27)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "廠  商"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 145)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 27)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "車  號"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label6.Location = New System.Drawing.Point(10, 207)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(84, 27)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "產  品"
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.MonthCalendar1.Location = New System.Drawing.Point(10, 87)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 7
        '
        'txtDate
        '
        Me.txtDate.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.txtDate.Location = New System.Drawing.Point(10, 39)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(220, 36)
        Me.txtDate.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.txtDate)
        Me.GroupBox1.Controls.Add(Me.MonthCalendar1)
        Me.GroupBox1.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 9)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(247, 261)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "日  期"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.cmbProduct)
        Me.GroupBox2.Controls.Add(Me.cmbCarNo)
        Me.GroupBox2.Controls.Add(Me.cmbManufacturer)
        Me.GroupBox2.Controls.Add(Me.cmbCustomer)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Location = New System.Drawing.Point(270, 18)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(330, 252)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        '
        'cmbProduct
        '
        Me.cmbProduct.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cmbProduct.FormattingEnabled = True
        Me.cmbProduct.Location = New System.Drawing.Point(100, 207)
        Me.cmbProduct.Name = "cmbProduct"
        Me.cmbProduct.Size = New System.Drawing.Size(217, 32)
        Me.cmbProduct.TabIndex = 9
        '
        'cmbCarNo
        '
        Me.cmbCarNo.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cmbCarNo.FormattingEnabled = True
        Me.cmbCarNo.Location = New System.Drawing.Point(100, 145)
        Me.cmbCarNo.Name = "cmbCarNo"
        Me.cmbCarNo.Size = New System.Drawing.Size(217, 32)
        Me.cmbCarNo.TabIndex = 8
        '
        'cmbManufacturer
        '
        Me.cmbManufacturer.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cmbManufacturer.FormattingEnabled = True
        Me.cmbManufacturer.Location = New System.Drawing.Point(100, 83)
        Me.cmbManufacturer.Name = "cmbManufacturer"
        Me.cmbManufacturer.Size = New System.Drawing.Size(217, 32)
        Me.cmbManufacturer.TabIndex = 7
        '
        'cmbCustomer
        '
        Me.cmbCustomer.Font = New System.Drawing.Font("新細明體", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cmbCustomer.FormattingEnabled = True
        Me.cmbCustomer.Location = New System.Drawing.Point(100, 21)
        Me.cmbCustomer.Name = "cmbCustomer"
        Me.cmbCustomer.Size = New System.Drawing.Size(217, 32)
        Me.cmbCustomer.TabIndex = 6
        '
        'btnQuery
        '
        Me.btnQuery.AutoSize = True
        Me.btnQuery.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnQuery.Location = New System.Drawing.Point(606, 150)
        Me.btnQuery.Name = "btnQuery"
        Me.btnQuery.Size = New System.Drawing.Size(155, 45)
        Me.btnQuery.TabIndex = 11
        Me.btnQuery.Text = "搜  尋"
        Me.btnQuery.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.AutoSize = True
        Me.btnClear.Font = New System.Drawing.Font("新細明體", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnClear.Location = New System.Drawing.Point(606, 213)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(155, 45)
        Me.btnClear.TabIndex = 12
        Me.btnClear.Text = "清  除"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'frmScaleQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 277)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnQuery)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmScaleQuery"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "查詢"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents MonthCalendar1 As MonthCalendar
    Friend WithEvents txtDate As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents cmbProduct As ComboBox
    Friend WithEvents cmbCarNo As ComboBox
    Friend WithEvents cmbManufacturer As ComboBox
    Friend WithEvents cmbCustomer As ComboBox
    Friend WithEvents btnQuery As Button
    Friend WithEvents btnClear As Button
End Class
