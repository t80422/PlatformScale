Module Utility
    ''' <summary>
    ''' 設定DataGridView的樣式屬性
    ''' </summary>
    ''' <param name="ctrl"></param>
    Public Sub SetDataGridViewStyle(ctrl As Control)
        For Each dgv In GetControlInParent(Of DataGridView)(ctrl)
            With dgv
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .ColumnHeadersDefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .DefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(224, 224, 224)
                .EnableHeadersVisualStyles = False
                .ColumnHeadersDefaultCellStyle.BackColor = Color.MediumTurquoise
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End With
        Next
    End Sub

    ''' <summary>
    ''' 取得指定控制項內所有的目標控制項
    ''' </summary>
    ''' <typeparam name="T">目標控制項</typeparam>
    ''' <param name="parent">指定控制項</param>
    ''' <returns></returns>
    Public Function GetControlInParent(Of T As Control)(parent As Control) As List(Of T)
        Dim lst As New List(Of T)
        If parent.Controls.Count > 0 Then
            For Each ctrl In parent.Controls
                If TypeOf ctrl Is T Then lst.Add(ctrl)
                lst.AddRange(GetControlInParent(Of T)(ctrl))
            Next
        End If
        Return lst
    End Function

    ''' <summary>
    ''' 清空指定控制項內的控制項
    ''' </summary>
    ''' <param name="control">容器型控制項</param>
    Public Sub ClearControl(control As Control)
        For Each ctrl As Control In control.Controls
            If TypeOf ctrl Is GroupBox Then
                ClearControl(ctrl)
            ElseIf TypeOf ctrl Is TabControl Then
                For Each tp As TabPage In CType(ctrl, TabControl).Controls
                    ClearControl(tp)
                Next
            End If
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Checked = False
            ElseIf TypeOf ctrl Is RadioButton Then
                CType(ctrl, RadioButton).Checked = False
            ElseIf TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).SelectedIndex = -1
            End If
        Next
    End Sub
    ''' <summary>
    ''' 清空指定控制項內的控制項
    ''' </summary>
    ''' <param name="control">控制項集合</param>
    Public Sub ClearControl(control As IEnumerable(Of Control))
        For Each ctrl As Control In control
            If TypeOf ctrl Is GroupBox Then
                ClearControl(ctrl)
            ElseIf TypeOf ctrl Is TabControl Then
                For Each tp As TabPage In CType(ctrl, TabControl).Controls
                    ClearControl(tp)
                Next
            End If
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Checked = False
            ElseIf TypeOf ctrl Is RadioButton Then
                CType(ctrl, RadioButton).Checked = False
            ElseIf TypeOf ctrl Is ComboBox Then
                Dim cmb As ComboBox = ctrl
                cmb.SelectedIndex = -1
                cmb.Text = ""
            End If
        Next
    End Sub

    ''' <summary>
    ''' 將取得的資料傳至各控制項(控制項的Tag必須寫上表格欄位名稱)
    ''' </summary>
    ''' <param name="ctrls">父容器</param>
    ''' <param name="row"></param>
    Public Sub GetDataToControls(ctrls As Control, row As Object)
        For Each ctrl In ctrls.Controls.Cast(Of Control).Where(Function(c) c.Tag IsNot Nothing)
            Dim column As String
            If ctrl.Tag.Contains("[") OrElse ctrl.Tag.Contains("]") Then
                column = ctrl.Tag.Replace("[", "").Replace("]", "")
            Else
                column = ctrl.Tag
            End If

            Dim value = GetCellData(row, column)
            Select Case ctrl.GetType.Name
                Case "TextBox"
                    ctrl.Text = value
                Case "DateTimePicker"
                    Dim dtp As DateTimePicker = ctrl
                    dtp.Value = value
                Case "ComboBox"
                    Dim cmb As ComboBox = ctrl
                    cmb.SelectedIndex = cmb.FindStringExact(value)
                Case "GroupBox"
                    Dim grp As GroupBox = ctrl
                    For Each c In grp.Controls
                        If TypeOf c Is CheckBox Then
                            Dim chk As CheckBox = c
                            chk.Checked = value = chk.Text
                        ElseIf TypeOf c Is RadioButton Then
                            Dim rdo As RadioButton = c
                            rdo.Checked = value = rdo.Text
                        End If
                    Next
                    GetDataToControls(ctrl, row)
                Case Else
            End Select
        Next
    End Sub

    Public Function GetCellData(row As Object, colName As String) As String
        Select Case row.GetType.Name
            Case "DataRow"
                Dim r As DataRow = row
                Return r(colName).ToString
            Case "DataGridViewRow"
                Dim r As DataGridViewRow = row
                Return r.Cells(colName).Value.ToString
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' 檢查TextBox裡是否為正數
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    Public Function CheckPositiveNumber(txt As TextBox) As Boolean
        If Not IsNumeric(txt.Text) Then
            MsgBox(txt.Tag + " 不為數字!")
            txt.Focus()
            Return False
        End If
        If Val(txt.Text) < 0 Then
            MsgBox(txt.Tag + " 不能為負數!")
            txt.Focus()
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 取得Table所有資料
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Public Function GetTableAllData(tableName As String) As String
        Select Case tableName
            Case "廠商資料表"
                Return "SELECT * FROM 廠商資料表"
            Case "客戶資料表"
                Return "SELECT * FROM 客戶資料表"
            Case "車籍資料表"
                Return "SELECT * FROM 車籍資料表"
            Case "產品資料表"
                Return "SELECT 代號, 品名, 每米噸數, 噸數設定 FROM 產品資料表"
            Case "過磅資料表"
                Return "SELECT * FROM 過磅資料表"
            Case "二次過磅暫存資料表"
                Return $"SELECT * FROM 二次過磅暫存資料表 WHERE 過磅日期 = '{Now:yyyy/MM/dd}' ORDER BY 過磅時間"
            Case "密碼資料表"
                Return "SELECT * FROM 密碼資料表"
        End Select
        Return ""
    End Function

    ''' <summary>
    ''' 檢查必填欄位
    ''' </summary>
    ''' <param name="required">填入key:Table欄位 value:控制項</param>
    ''' <returns></returns>
    Public Function CheckRequiredCol(required As Dictionary(Of String, Object)) As Boolean
        For Each kvp In required
            If String.IsNullOrWhiteSpace(kvp.Value.Text) Then
                MsgBox(kvp.Key + " 不能空白")
                kvp.Value.Focus()
                Return False
            End If
        Next
        Return True
    End Function

End Module
