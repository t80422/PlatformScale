Imports System.IO
Imports System.Configuration
Imports System.IO.Ports

Public Class frmMain
    Public permissions As Integer '權限等級
    Public user As String '使用者
    Private serialPortA As SerialPort
    Private serialPortB As SerialPort
    Private nowScale As String
    Private portClose As Boolean
    Private spAClose As Integer 'Port停止傳訊息就表示關閉了,此變數用來紀錄逾時
    Private spBClose As Integer 'Port停止傳訊息就表示關閉了,此變數用來紀錄逾時

    Private Enum enumWho
        客戶
        廠商
    End Enum

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim folderPath As String = Path.Combine(Application.StartupPath, "Report")
        '檢查資料夾是否存在
        If Not Directory.Exists(folderPath) Then Directory.CreateDirectory(folderPath)
        InitDataGrid()
        Init車籍()
        Init過磅()
        '初始化權限設定的cmb權限
        With cmb權限.Items
            .Add("1")
            .Add("2")
            .Add("3")
        End With
        '讀取系統設定-遠端備份
        lblRemote.Text = SelectTable($"SELECT IP FROM 遠端備份資料表").Rows(0)("IP")
        '初始化 系統設定-Port設定
        dgvPort.DataSource = SelectTable("SELECT * FROM 通訊埠口資料表")
    End Sub

    Private Sub InitDataGrid()
        SetDataGridViewStyle(Me)
        dgv廠商.DataSource = SelectTable(GetTableAllData("廠商資料表"))
        dgv客戶.DataSource = SelectTable(GetTableAllData("客戶資料表"))
        dgv車籍.DataSource = SelectTable(GetTableAllData("車籍資料表"))
        dgv貨品.DataSource = SelectTable(GetTableAllData("產品資料表"))
        dgv權限.DataSource = SelectTable(GetTableAllData("密碼資料表"))
    End Sub

    Private Sub Init車籍()
        lst廠商.DataSource = SelectTable("SELECT DISTINCT 簡稱 FROM 廠商資料表")
        lst廠商.DisplayMember = "簡稱"
        lst客戶.DataSource = SelectTable("SELECT DISTINCT 簡稱 FROM 客戶資料表")
        lst客戶.DisplayMember = "簡稱"
    End Sub

    Private Sub Init過磅()
        '設定客戶
        SetCmbCliManu(enumWho.客戶)
        SetComPort()
        SetScaleSign()
        ScaleSwitch("A")
        btnClear_過磅.PerformClick()
    End Sub

    ''' <summary>
    ''' 設定電子磅秤警示燈
    ''' </summary>
    Private Sub SetScaleSign()
        Dim aCircle As New Drawing2D.GraphicsPath
        aCircle.AddEllipse(New RectangleF(0, 0, 50, 50))
        lblA.Size = New Size(50, 50)
        lblA.Region = New Region(aCircle)
        lblB.Size = New Size(50, 50)
        lblB.Region = New Region(aCircle)
    End Sub

    ''' <summary>
    ''' 設定Port參數
    ''' </summary>
    Private Sub SetComPort()
        Dim row = SelectTable("SELECT * FROM 通訊埠口資料表").Rows(0)
        serialPortA = New SerialPort("COM" + row("埠口1").ToString, 9600, Parity.None, 7, StopBits.One)
        AddHandler serialPortA.DataReceived, AddressOf SerialPortA_DataReceived
        Try
            If serialPortA.IsOpen Then serialPortA.Close()
            serialPortA.Open()
            portClose = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        serialPortB = New SerialPort("COM" + row("埠口2").ToString, 9600, Parity.None, 7, StopBits.One)
        AddHandler serialPortB.DataReceived, AddressOf SerialPortB_DataReceived
        Try
            If serialPortB.IsOpen Then serialPortB.Close()
            serialPortB.Open()
            portClose = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 磅秤警示閃爍
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub tmrScale_Tick(sender As Object, e As EventArgs) Handles tmrScale.Tick
        Select Case nowScale
            Case "A"
                With lblA
                    If .BackColor = Color.White Then
                        .BackColor = Color.Red
                    Else
                        .BackColor = Color.White
                    End If
                End With
            Case "B"
                With lblB
                    If .BackColor = Color.White Then
                        .BackColor = Color.Red
                    Else
                        .BackColor = Color.White
                    End If
                End With
        End Select
    End Sub

    ''' <summary>
    ''' 設定切換表頭
    ''' </summary>
    ''' <param name="AorB"></param>
    Private Sub ScaleSwitch(AorB As String)
        Select Case AorB
            Case "A"
                nowScale = AorB
                lblB.BackColor = Color.White
            Case "B"
                nowScale = AorB
                lblA.BackColor = Color.White
        End Select
        tmrScale.Enabled = True
    End Sub

    '過磅作業-選項-進貨
    Private Sub rdoPurchase_CheckedChanged(sender As RadioButton, e As EventArgs) Handles rdoPurchase.CheckedChanged
        If sender.Checked Then
            lblCliManu.Text = "廠    商"
            SetCmbCliManu(enumWho.廠商)
        End If
    End Sub

    '過磅作業-選項-出貨
    Private Sub rdoShipment_CheckedChanged(sender As Object, e As EventArgs) Handles rdoShipment.CheckedChanged
        If sender.Checked Then
            lblCliManu.Text = "客    戶"
            SetCmbCliManu(enumWho.客戶)
        End If
    End Sub

    '過磅作業-車號
    Private Sub cmbCarNo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbCarNo.SelectionChangeCommitted
        '將所選車輛的"空重"傳至"空車重量"
        Dim row As DataRowView = cmbCarNo.SelectedItem
        If IsDBNull(row("空重")) OrElse String.IsNullOrEmpty(row("空重")) Then
            txtLoudTime_Empty.Clear()
            txtEmptyCar.Clear()
        Else
            txtEmptyCar.Text = row("空重")
            txtLoudTime_Empty.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
        End If
        txtCarCount.Text = GetCarCount(row("車號"))
    End Sub

    ''' <summary>
    ''' 計算當日車次
    ''' </summary>
    ''' <param name="carNo"></param>
    ''' <returns></returns>
    Private Function GetCarCount(carNo As String) As Integer
        Dim carData = SelectTable($"SELECT COUNT(車牌號碼) AS coun FROM 過磅資料表 WHERE 車牌號碼 = '{carNo}' AND 過磅日期 = '{Now:yyyy/MM/dd}'").Rows
        Dim carCount As Integer
        If carData.Count > 0 Then
            carCount = carData(0)("coun")
        Else
            carCount = 0
        End If
        Return carCount
    End Function

    '過磅作業-載入空重
    Private Sub btnLoadEmpty_Click(sender As Object, e As EventArgs) Handles btnLoadEmpty.Click
        LoadWeight(txtEmptyCar, txtLoudTime_Empty)
    End Sub

    '過磅作業-載入總重
    Private Sub btnLoadTotal_Click(sender As Object, e As EventArgs) Handles btnLoadTotal.Click
        LoadWeight(txtTotalWeight, txtLoudTime_Total)
    End Sub

    Private Sub LoadWeight(weight As TextBox, time As TextBox)
        Dim value = GetWeight()
        weight.Text = value
        time.Text = Date.Parse(lblTime.Text).ToString("HH:mm")

    End Sub

    ''' <summary>
    ''' 取得當前磅秤重量
    ''' </summary>
    ''' <returns></returns>
    Private Function GetWeight() As Double
        Dim value = ""
        Select Case nowScale
            Case "A"
                value = lblAValue.Text
            Case "B"
                value = lblBValue.Text
        End Select

        If Not IsNumeric(value) Then
            MsgBox($"電子磅秤 {nowScale} 無法讀取有效重量值！")
            Return 0
        Else
            Return Format(Val(value) / 1000, "###.#0")
        End If
    End Function

    '擷取A磅秤資料
    Private Sub SerialPortA_DataReceived(sender As SerialPort, e As SerialDataReceivedEventArgs)
        If portClose Then
            serialPortA.Close()
            Exit Sub
        End If
        If sender.BytesToRead >= 19 Then
            Dim txt = sender.ReadExisting()
            Try
                Invoke(Sub()
                           Dim weight = Mid(txt, 7, 8).Replace(" ", "")
                           If Not IsNumeric(weight) Then Exit Sub
                           lblAValue.Text = Double.Parse(weight).ToString()
                       End Sub)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            spAClose = 0
        End If
    End Sub

    '擷取B磅秤資料
    Private Sub SerialPortB_DataReceived(sender As SerialPort, e As SerialDataReceivedEventArgs)
        If portClose Then
            serialPortB.Close()
            Exit Sub
        End If
        If sender.BytesToRead >= 19 Then
            Dim txt = sender.ReadExisting()
            Try
                Invoke(Sub()
                           Dim weight = Mid(txt, 7, 8).Replace(" ", "")
                           If Not IsNumeric(weight) Then Exit Sub
                           lblBValue.Text = Double.Parse(weight).ToString()
                       End Sub)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            spBClose = 0
        End If
    End Sub

    '過磅作業-產品
    Private Sub cmbProduct_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbProduct.SelectionChangeCommitted
        Dim cmb = CType(sender, ComboBox)
        Dim row As DataRowView = cmb.SelectedItem
        txtTPM.Text = row("每米噸數")
    End Sub

    ''' <summary>
    ''' 設定客戶、廠商(Manufacturers) 下拉選單
    ''' </summary>
    ''' <param name="value"></param>
    Private Sub SetCmbCliManu(value As enumWho)
        Dim table = ""
        Select Case value
            Case enumWho.客戶
                table = "客戶資料表"
            Case enumWho.廠商
                table = "廠商資料表"
        End Select
        With cmbCliManu
            .DataSource = SelectTable($"SELECT 簡稱 FROM {table}")
            .DisplayMember = "簡稱"
            .SelectedIndex = -1
        End With
    End Sub

    Private Sub tmr過磅_Tick(sender As Object, e As EventArgs) Handles tmr過磅.Tick
        lblTime.Text = Now.ToString("HH:mm:ss")
        spAClose += 1
        spBClose += 1
        If spAClose >= 5 Then lblAValue.Text = "-----"
        If spBClose >= 5 Then lblBValue.Text = "-----"
    End Sub

    ''' <summary>
    ''' 取最新磅單序號
    ''' </summary>
    Private Function GetNewRecpNo() As String
        Dim d = dtp過磅.Value
        Dim dt = SelectTable($"SELECT 磅單序號 FROM 過磅資料表 WHERE 過磅日期 = '{d:yyyy/MM/dd}' " +
                              "UNION " +
                             $"SELECT 磅單序號 FROM 二次過磅暫存資料表 WHERE 過磅日期 = '{d:yyyy/MM/dd}' " +
                              "ORDER BY 磅單序號 DESC")
        Dim num As String
        If dt.Rows.Count = 0 Then
            num = d.ToString("yyyyMMdd001")
        Else
            num = dt.Rows(0).Field(Of String)("磅單序號") + 1
        End If
        Return num
    End Function

    '過磅作業-切換表頭
    Private Sub btnSwitchScale_Click(sender As Object, e As EventArgs) Handles btnSwitchScale.Click
        Select Case nowScale
            Case "A"
                ScaleSwitch("B")
            Case "B"
                ScaleSwitch("A")
        End Select

    End Sub

    '過磅作業-列印
    Private Sub btnPrint_過磅_Click(sender As Object, e As EventArgs) Handles btnPrint_過磅.Click
        If dgv過磅.Rows.Count = 0 Then
            MsgBox("無任何記錄可供列印！")
            Exit Sub
        End If
        If String.IsNullOrEmpty(txtRcepNo.Text) Then
            MsgBox("請選擇列印磅單")
            Exit Sub
        End If

    End Sub

    '過磅作業-淨重、每米頓數-改變就計算總米數
    Private Sub txtNetWeight_TextChanged(sender As Object, e As EventArgs) Handles txtNetWeight.TextChanged, txtTPM.TextChanged
        If String.IsNullOrEmpty(txtTPM.Text) OrElse Not IsNumeric(txtTPM.Text) OrElse String.IsNullOrEmpty(txtNetWeight.Text) Then Exit Sub
        '若是負數則顯示0
        If txtNetWeight.Text < 0 Then txtNetWeight.Text = 0

        If txtTPM.Text = 0 Then
            txtMeter.Clear()
        Else
            txtMeter.Text = (Double.Parse(txtNetWeight.Text) / Double.Parse(txtTPM.Text)).ToString(grpDecimal.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text)
        End If
    End Sub

    '過磅作業-空車重量、總重-總重與空車重量有值才計算,改變就動態計算淨重
    Private Sub txtEmptyCar_TextChanged(sender As Object, e As EventArgs) Handles txtEmptyCar.TextChanged, txtTotalWeight.TextChanged
        If String.IsNullOrEmpty(txtTotalWeight.Text) OrElse String.IsNullOrEmpty(txtEmptyCar.Text) OrElse Not IsNumeric(txtTotalWeight.Text) OrElse Not IsNumeric(txtEmptyCar.Text) Then
            txtNetWeight.Clear()
            Exit Sub
        End If
        txtNetWeight.Text = (Double.Parse(txtTotalWeight.Text) - Double.Parse(txtEmptyCar.Text)).ToString("0.00")
    End Sub

    ''過磅作業-空車重量
    'Private Sub txtEmptyCar_Leave(sender As Object, e As EventArgs) Handles txtEmptyCar.Leave
    '    '如果欄位為空,就清除載入時間
    '    If String.IsNullOrWhiteSpace(txtEmptyCar.Text) Then
    '        txtLoudTime_Empty.Clear()
    '        Exit Sub
    '    End If
    '    '檢查是否為正數,是就填入時間,防止使用者手動輸入完資料後沒按Enter
    '    If CheckPositiveNumber(txtEmptyCar) AndAlso String.IsNullOrWhiteSpace(txtLoudTime_Empty.Text) Then MsgBox("請輸入Enter")
    'End Sub

    ''過磅作業-總重
    'Private Sub txtTotalWeight_Leave(sender As Object, e As EventArgs) Handles txtTotalWeight.Leave
    '    '如果欄位為空,就清除載入時間
    '    If String.IsNullOrWhiteSpace(txtTotalWeight.Text) Then
    '        txtLoudTime_Total.Clear()
    '        Exit Sub
    '    End If
    '    '檢查是否為正數,是就填入時間,防止使用者手動輸入完資料後沒按Enter
    '    If CheckPositiveNumber(txtTotalWeight) AndAlso String.IsNullOrWhiteSpace(txtLoudTime_Total.Text) Then MsgBox("請輸入Enter")
    'End Sub

    ''過磅作業-空車重量-按 Enter 代入 載入時間
    'Private Sub txtEmptyCar_KeyPress(sender As TextBox, e As KeyPressEventArgs) Handles txtEmptyCar.KeyPress
    '    If sender.ReadOnly = False AndAlso e.KeyChar = vbCr Then
    '        txtLoudTime_Empty.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
    '    End If
    'End Sub

    ''過磅作業-總重-按 Enter 代入 載入時間
    'Private Sub txtTotalWeight_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTotalWeight.KeyPress
    '    If sender.ReadOnly = False AndAlso e.KeyChar = vbCr Then
    '        txtLoudTime_Total.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
    '    End If
    'End Sub

    '過磅作業-總米數單位.00
    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        If RadioButton13.Checked AndAlso Not String.IsNullOrEmpty(txtMeter.Text) Then
            txtMeter.Text = Double.Parse(txtMeter.Text).ToString("0.00")
        End If
    End Sub

    '過磅作業-總米數單位.000
    Private Sub RadioButton12_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton12.CheckedChanged
        If RadioButton12.Checked AndAlso Not String.IsNullOrEmpty(txtMeter.Text) Then
            txtMeter.Text = Double.Parse(txtMeter.Text).ToString("0.000")
        End If
    End Sub

    '過磅作業-重量輸入-自動
    Private Sub rdoAuto_CheckedChanged(sender As RadioButton, e As EventArgs) Handles rdoAuto.CheckedChanged
        If sender.Checked Then
            txtEmptyCar.ReadOnly = True
            txtTotalWeight.ReadOnly = True
        End If
    End Sub

    '過磅作業-重量輸入-手動
    Private Sub rdoManual_CheckedChanged(sender As RadioButton, e As EventArgs) Handles rdoManual.CheckedChanged
        If sender.Checked Then
            txtEmptyCar.ReadOnly = False
            txtTotalWeight.ReadOnly = False
        End If
    End Sub

    '貨品資料-噸數設定-不設定
    Private Sub rdoTUnset_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTUnset.CheckedChanged
        If rdoTUnset.Checked Then
            txtTM.Enabled = False
            txtTM.Text = 0
        End If
    End Sub

    '貨品資料-噸數設定-設定
    Private Sub rdoTSet_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTSet.CheckedChanged
        txtTM.Enabled = rdoTSet.Checked
    End Sub

    '清除-過磅作業
    Private Sub btnClear_過磅_Click(sender As Button, e As EventArgs) Handles btnClear_過磅.Click
        ClearControl(sender.Parent.Controls.OfType(Of Control).Where(Function(ctrl) TypeOf ctrl IsNot GroupBox))
        '刷新當日在場內車輛列表
        With dgv二次過磅
            .DataSource = SelectTable(GetTableAllData("二次過磅暫存資料表"))
            .Columns("空重載入時間").DefaultCellStyle.Format = "HH:mm"
            .Columns("總重載入時間").DefaultCellStyle.Format = "HH:mm"
            .Columns("磅單序號").Visible = False
            .Columns("過磅種類").Visible = False
            .Columns("過磅日期").Visible = False
            .Columns("工程名稱").Visible = False
            .Columns("工程代號").Visible = False
            .Columns("載運地點").Visible = False
            .Columns("承辦人").Visible = False
            .Columns("產品代號").Visible = False
            .Columns("淨重").Visible = False
            .Columns("每米噸數").Visible = False
            .Columns("米數").Visible = False
            .Columns("備註").Visible = False
            .Columns("過磅時間").Visible = False
        End With
        With dgv過磅
            .DataSource = SelectTable(GetTableAllData("過磅資料表") + $" WHERE 過磅日期 = '{Now:yyyy/MM/dd}' ORDER BY 磅單序號 DESC")
            .Columns("空重載入時間").DefaultCellStyle.Format = "HH:mm"
            .Columns("總重載入時間").DefaultCellStyle.Format = "HH:mm"
            .Columns("過磅種類").Visible = False
            .Columns("工程名稱").Visible = False
            .Columns("工程代號").Visible = False
            .Columns("載運地點").Visible = False
            .Columns("產品代號").Visible = False
            .Columns("備註").Visible = False
            .Columns("全銜").Visible = False
            .Columns("單價").Visible = False
            .Columns("總價").Visible = False
        End With
        '設定cmb產品
        With cmbProduct
            .DataSource = SelectTable($"SELECT * FROM 產品資料表")
            .DisplayMember = "品名"
            .SelectedIndex = -1
        End With

        cmbCliManu.Enabled = False
        cmbCarNo.Enabled = False
        cmbProduct.Enabled = False
        txtTPM.ReadOnly = True
    End Sub

    '清除-系統設定-權限設定
    Private Sub btnClear_權限_Click(sender As Object, e As EventArgs) Handles btnClear_權限.Click
        ClearControl(grp權限)
        dgv權限.DataSource = SelectTable(GetTableAllData("密碼資料表"))
        cmb權限.SelectedIndex = -1
    End Sub

    '新增-過磅作業
    Private Sub btnInsert_過磅_Click(sender As Object, e As EventArgs) Handles btnInsert_過磅.Click
        btnClear_過磅.PerformClick()
        txtRcepNo.Text = GetNewRecpNo()
        txtUser.Text = user
        cmbCliManu.Enabled = True
        cmbCarNo.Enabled = True
        cmbProduct.Enabled = True
        txtTPM.ReadOnly = False
    End Sub

    '新增-廠商資料,客戶資料
    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert_廠商.Click, btnInsert_客戶.Click
        Dim tp = CType(sender, Button).Parent
        Dim table As String = ""
        Dim btn As Button = Nothing
        Dim lst As New List(Of Object)
        Dim dic As New Dictionary(Of String, Object)
        Select Case tp.Text
            Case "廠商資料"
                table = "廠商資料表"
                btn = btnClear_廠商
                dic.Add("全銜", txtName_廠商)
                dic.Add("代號", txtNo_廠商)
                dic.Add("簡稱", txtAka_廠商)
                lst.Add(txtNo_廠商)
            Case "客戶資料"
                table = "客戶資料表"
                btn = btnClear_客戶
                dic.Add("全銜", txtName_客戶)
                dic.Add("代號", txtNo_客戶)
                dic.Add("簡稱", txtAka_客戶)
                lst.Add(txtNo_客戶)
        End Select
        If CheckInsert(sender, dic, lst, table) Then
            btn.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '新增-貨品資料
    Private Sub btnInsert_貨品_Click(sender As Object, e As EventArgs) Handles btnInsert_貨品.Click
        If Not Check貨品() Then Exit Sub
        '取得各欄位的值
        Dim dic As New Dictionary(Of String, Object)
        tp貨品.Controls.OfType(Of TextBox).Where(Function(txt) txt.Tag <> "每米噸數").ToList.ForEach(Sub(txt) dic.Add(txt.Tag.ToString, txt.Text))
        If rdoTSet.Checked Then
            dic.Add("噸數設定", "1")
            dic.Add(txtTM.Tag, txtTM.Text)
        End If
        If InserTable("產品資料表", dic) Then
            btnClear_貨品.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '新增-系統設定-權限設定
    Private Sub btnInsert_權限_Click(sender As Object, e As EventArgs) Handles btnInsert_權限.Click
        Dim dic As New Dictionary(Of String, Object)
        grp權限.Controls.OfType(Of Control).Where(Function(txt) txt.Tag <> "").ToList.ForEach(Sub(ctrl) dic.Add(ctrl.Tag.ToString, ctrl.Text))
        If InserTable("密碼資料表", dic) Then
            btnClear_權限.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '新增-車籍資料
    Private Sub txtInsert_車籍_Click(sender As Object, e As EventArgs) Handles txtInsert_車籍.Click
        Dim dic As New Dictionary(Of String, Object) From {
            {"車號", txtNo_車籍},
            {"車主", txt車主}
        }
        Dim lst As New List(Of Object) From {txtNo_車籍}
        If CheckInsert(sender, dic, lst, "車籍資料表") Then
            btnClear_車籍.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    'dgv點擊-過磅作業-場內車輛列表(二次過磅)、過磅
    Private Sub dgv過磅_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv二次過磅.CellMouseClick, dgv過磅.CellMouseClick
        Dim dgv As DataGridView = sender
        If dgv.SelectedRows.Count < 1 Then Exit Sub
        ClearControl(dgv.Parent.Controls.OfType(Of Control).Where(Function(ctrl) TypeOf ctrl IsNot GroupBox))
        Dim selectRow = dgv.SelectedRows(0)
        GetDataToControls(tp過磅, selectRow)
        '因為會切換到選項的進出貨,就會觸發cmb重置,所以要再抓一次
        Dim value = GetCellData(selectRow, "客戶/廠商")
        cmbCliManu.SelectedIndex = cmbCliManu.FindStringExact(value)
        Dim carNo = GetCellData(selectRow, "車牌號碼")
        cmbCarNo.SelectedIndex = cmbCarNo.FindStringExact(carNo)
        If Not String.IsNullOrEmpty(txtLoudTime_Empty.Text) Then txtLoudTime_Empty.Text = Date.Parse(txtLoudTime_Empty.Text).ToString("HH:mm")
        If Not String.IsNullOrEmpty(txtLoudTime_Total.Text) Then txtLoudTime_Total.Text = Date.Parse(txtLoudTime_Total.Text).ToString("HH:mm")

        '計算當日車趟數
        txtCarCount.Text = GetCarCount(carNo)

        If permissions <> 1 Then
            cmbCliManu.Enabled = True
            cmbProduct.Enabled = True
            txtTPM.ReadOnly = False
            cmbCarNo.Enabled = True
        End If
    End Sub

    '車號隨客戶/廠商改變
    Private Sub cmbCliManu_TextChanged(sender As Object, e As EventArgs) Handles cmbCliManu.TextChanged
        cmbCarNo.DataSource = Nothing
        Dim row As DataRowView = cmbCliManu.SelectedItem
        If row IsNot Nothing Then
            '所選客戶/廠商 車號
            With cmbCarNo
                .DataSource = SelectTable($"SELECT * FROM 車籍資料表 WHERE 車主 = '{row("簡稱")}'")
                .DisplayMember = "車號"
                .SelectedIndex = -1
            End With
        End If

        '刷新空車重,載入時間
        txtEmptyCar.Clear()
        txtLoudTime_Empty.Clear()
    End Sub

    'dgv點擊-系統設定-權限設定
    Private Sub dgv權限_CellMouseClick(sender As DataGridView, e As DataGridViewCellMouseEventArgs) Handles dgv權限.CellMouseClick
        Dim dgv As DataGridView = sender
        If dgv.SelectedRows.Count < 1 Then Exit Sub
        ClearControl(grp權限)
        Dim selectRow = dgv.SelectedRows(0)
        GetDataToControls(grp權限, selectRow)
    End Sub

    '儲存-過磅作業
    Private Sub btnSave_過磅_Click(sender As Object, e As EventArgs) Handles btnSave_過磅.Click
        If String.IsNullOrEmpty(txtRcepNo.Text) Then
            MsgBox("請先 新增過磅")
            Exit Sub
        End If

        If Not String.IsNullOrEmpty(txtLoudTime_Empty.Text) And Not String.IsNullOrEmpty(txtLoudTime_Total.Text) Then
            If SelectTable($"SELECT 磅單序號 FROM 過磅資料表 WHERE {txtRcepNo.Tag} = '{txtRcepNo.Text}'").Rows.Count > 0 Then
                If MsgBox($"{txtRcepNo.Tag}:{txtRcepNo.Text} 已修改,是否覆蓋", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
            End If
            DeleteTable("過磅資料表", $"{txtRcepNo.Tag} = '{txtRcepNo.Text}'")
            If Insert過磅Data("過磅資料表") Then
                Dim dic As New Dictionary(Of String, String) From {{"車牌號碼", cmbCarNo.Text}}
                DeleteTable("二次過磅暫存資料表", dic)
            Else
                Exit Sub
            End If
            '二次過磅(對象為未登錄空車重)
        ElseIf Not String.IsNullOrEmpty(txtLoudTime_Empty.Text) OrElse Not String.IsNullOrEmpty(txtLoudTime_Total.Text) Then
            ''提醒輸入載入時間
            'If Not String.IsNullOrWhiteSpace(txtEmptyCar.Text) AndAlso Integer.Parse(txtEmptyCar.Text) > 0 AndAlso String.IsNullOrWhiteSpace(txtLoudTime_Empty.Text) Then
            '    MsgBox("請輸入空車載入時間,否則空車重量請輸入0")
            '    Exit Sub
            'End If
            'If Not String.IsNullOrWhiteSpace(txtTotalWeight.Text) AndAlso Integer.Parse(txtTotalWeight.Text) > 0 AndAlso String.IsNullOrWhiteSpace(txtLoudTime_Total.Text) Then
            '    MsgBox("請輸入總重載入時間,否則總重請輸入0")
            '    Exit Sub
            'End If

            Dim tempRows = SelectTable($"SELECT * FROM 二次過磅暫存資料表 WHERE {cmbCarNo.Tag} = '{cmbCarNo.Text}' AND 過磅日期 = '{Now:yyyy/MM/dd}'").Rows
            '同一台車不可能重複出現在廠內
            If tempRows.Count > 0 Then
                Dim rcepNo = tempRows(0).Field(Of String)("磅單序號")
                If rcepNo <> txtRcepNo.Text Then
                    MsgBox($"{cmbCarNo.Tag}:{cmbCarNo.Text} 重複輸入,請檢察 場內車輛列表")
                    Exit Sub
                ElseIf tempRows(0).Field(Of String)("磅單序號") = txtRcepNo.Text Then
                    Dim dic As New Dictionary(Of String, String) From {{"車牌號碼", cmbCarNo.Text}}
                    DeleteTable("二次過磅暫存資料表", dic)
                End If
            End If
            If Not Insert過磅Data("二次過磅暫存資料表") Then Exit Sub
        Else
            MsgBox("請先檢查重量貨載入時間")
            Exit Sub
        End If

        '臨時客戶/廠商,新增到資料表
        If cmbCliManu.SelectedIndex = -1 Then
            Dim table As String = ""
            Dim cm As enumWho
            '判斷是客戶還是廠商
            Select Case lblCliManu.Text
                Case "客    戶"
                    table = "客戶資料表"
                    cm = enumWho.客戶
                Case "廠    商"
                    table = "廠商資料表"
                    cm = enumWho.廠商
            End Select
            '檢查是否重複
            If SelectTable($"SELECT 簡稱 FROM {table} WHERE 簡稱 = '{cmbCliManu.Text}'").Rows.Count = 0 Then
                '取得代號
                Dim dtNo = SelectTable($"SELECT TOP 1 代號 FROM {table} ORDER BY 代號 DESC")
                Dim no As String
                If dtNo.Rows.Count > 0 Then
                    no = dtNo.Rows(0)("代號") + 1
                Else
                    no = 1
                End If
                '全銜、簡稱一樣
                Dim dic As New Dictionary(Of String, Object) From {
                    {"代號", no},
                    {"簡稱", cmbCliManu.Text},
                    {"全銜", cmbCliManu.Text}
                }
                InserTable(table, dic)
                SetCmbCliManu(cm)
                '對應臨時客戶/廠商新增時要刷新
                btnClear_Click(btnClear_客戶, e)
                btnClear_Click(btnClear_廠商, e)
            End If
        End If

        '臨時車號,新增到資料表
        If cmbCarNo.SelectedIndex = -1 Then
            '檢查是否重複
            If SelectTable($"SELECT 車號 FROM 車籍資料表 WHERE 車號 = '{cmbCarNo.Text}'").Rows.Count = 0 Then
                Dim dic As New Dictionary(Of String, Object) From {
                    {"車主", cmbCliManu.Text},
                    {"車號", cmbCarNo.Text}
                }
                InserTable("車籍資料表", dic)
            End If
            '對應臨時車號新增時要刷新
            btnClear_Click(btnClear_車籍, e)
        End If

        btnClear_過磅.PerformClick()
        MsgBox("儲存成功")
    End Sub

    Private Function Insert過磅Data(table As String) As Boolean
        Dim dicRequired As New Dictionary(Of String, Object) From {
            {"廠商/客戶", cmbCliManu},
            {"產品名稱", cmbProduct},
            {"車號", cmbCarNo}
        }
        If Not CheckRequiredCol(dicRequired) Then Return False
        Dim dicInsertData As New Dictionary(Of String, String)
        For Each ctrl In tp過磅.Controls.OfType(Of Control).Where(Function(ctrls) ctrls.Tag IsNot Nothing AndAlso ctrls.Text <> "")
            Select Case ctrl.GetType.Name
                Case "TextBox"
                    dicInsertData.Add(ctrl.Tag, ctrl.Text)
                Case "ComboBox"
                    dicInsertData.Add(ctrl.Tag, ctrl.Text)
                Case "DateTimePicker"
                    Dim dtp As DateTimePicker = ctrl
                    dicInsertData.Add(ctrl.Tag, dtp.Value.ToString("yyyy/MM/dd"))
                Case "Label"
                    dicInsertData.Add(ctrl.Tag, ctrl.Text)
            End Select
        Next
        dicInsertData.Add(grpInOut.Tag, grpInOut.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text)
        If InserTable(table, dicInsertData) Then
            Return True
        Else
            Return False
        End If
    End Function

    '儲存-貨品資料
    Private Sub btnModify_貨品_Click(sender As Object, e As EventArgs) Handles btnModify_貨品.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法修改")
            Exit Sub
        End If
        Dim condition() As String = {}
        If Not Check貨品() Then Exit Sub
        '取得欄位資料
        Dim dic As New Dictionary(Of String, String)
        condition = {"代號", txtNo_貨品.Text}
        tp貨品.Controls.OfType(Of TextBox).Where(Function(txt) txt.Tag <> condition(0)).ToList.ForEach(Sub(txt) dic.Add(txt.Tag.ToString, txt.Text))
        If rdoTSet.Checked Then
            dic.Add("噸數設定", "1")
        Else
            dic.Add("噸數設定", "0")
        End If
        If Not UpdateTable("產品資料表", dic, $"{condition(0)} = '{condition(1)}'") Then Exit Sub
        btnClear_貨品.PerformClick()
        MsgBox("修改成功")
    End Sub

    ''' <summary>
    ''' 新增、修改時檢查貨品 必填欄位、資料型態等
    ''' </summary>
    ''' <returns></returns>
    Private Function Check貨品() As Boolean
        '檢查每米噸數在選擇設定時有沒有填寫
        If rdoTSet.Checked Then
            If String.IsNullOrEmpty(txtTM.Text) Then
                MsgBox(txtTM.Tag + " 未填寫")
                txtTM.Focus()
                Return False
            End If
            If Not CheckPositiveNumber(txtTM) Then Return False
        End If

        Dim dicRequired As New Dictionary(Of String, Object) From {
            {txtNo_貨品.Tag, txtNo_貨品},
            {txtName_貨品.Tag, txtName_貨品}
        }
        If Not CheckRequiredCol(dicRequired) Then Return False
        Return True
    End Function

    '儲存-系統設定-Port設定
    Private Sub btnSave_Port_Click(sender As Object, e As EventArgs) Handles btnSave_Port.Click
        If String.IsNullOrWhiteSpace(txtPortA.Text) And String.IsNullOrWhiteSpace(txtPortB.Text) Then Exit Sub

        '輸入防呆
        If Not String.IsNullOrWhiteSpace(txtPortA.Text) AndAlso Not CheckPositiveNumber(txtPortA) Then Exit Sub
        If Not String.IsNullOrWhiteSpace(txtPortB.Text) AndAlso Not CheckPositiveNumber(txtPortB) Then Exit Sub

        '檢查是否兩個Port設定一樣
        Dim row = SelectTable("SELECT * FROM 通訊埠口資料表").Rows(0)
        If Not String.IsNullOrWhiteSpace(txtPortA.Text) AndAlso txtPortA.Text = row("埠口2") Then
            MsgBox("不能與 埠口2 設定一樣")
            Exit Sub
        End If
        If Not String.IsNullOrWhiteSpace(txtPortB.Text) AndAlso txtPortB.Text = row("埠口1") Then
            MsgBox("不能與 埠口1 設定一樣")
            Exit Sub
        End If

        Dim dic As New Dictionary(Of String, String)
        grpPort.Controls.OfType(Of TextBox).Where(Function(txt) Not String.IsNullOrWhiteSpace(txt.Text)).ToList.ForEach(Sub(t) dic.Add(t.Tag.ToString, t.Text))
        If UpdateTable("通訊埠口資料表", dic, "1=1") Then
            dgvPort.DataSource = SelectTable("SELECT * FROM 通訊埠口資料表")
            txtPortA.Clear()
            txtPortB.Clear()
            MsgBox("儲存成功")
        End If
    End Sub

    '儲存-系統設定-遠端備份
    Private Sub btnRemote_Click(sender As Object, e As EventArgs) Handles btnRemote.Click
        txtRemote.Enabled = False
        If Not String.IsNullOrWhiteSpace(txtRemote.Text) Then
            Dim dic As New Dictionary(Of String, String) From {{txtRemote.Tag, txtRemote.Text}}
            If UpdateTable("遠端備份資料表", dic, "1=1") Then
                lblRemote.Text = SelectTable($"SELECT IP FROM 遠端備份資料表").Rows(0)("IP")
                txtRemote.Clear()
                MsgBox("儲存成功")
            Else
                Exit Sub
            End If
        End If
        lblBackUping.Visible = True
        If lblRemote.Text.EndsWith("\") Then lblRemote.Text = lblRemote.Text.Remove(lblRemote.Text.Length - 1)
        Dim dbPath = lblRemote.Text + "\db4UGWS.mdb"
        Dim logPath = lblRemote.Text + "\DBBackUp.log"
        Try
            File.Copy(Application.StartupPath + "\db4UGWS.mdb", dbPath, True)
            File.AppendAllText(logPath, Now.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf)
            MsgBox("備份成功")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        txtRemote.Enabled = True
        lblBackUping.Visible = False
    End Sub

    '儲存-系統設定-權限設定
    Private Sub btnSave_權限_Click(sender As Object, e As EventArgs) Handles btnSave_權限.Click
        Dim dic As New Dictionary(Of String, String) From {
            {txtPsw_權限.Tag.ToString, txtPsw_權限.Text},
            {cmb權限.Tag.ToString, cmb權限.Text}
        }
        If UpdateTable("密碼資料表", dic, $"{txtName_權限.Tag} = '{txtName_權限.Text}'") Then
            btnClear_權限.PerformClick()
            MsgBox("儲存成功")
        End If
    End Sub

    '儲存-車籍資料
    Private Sub btnModify_車籍_Click(sender As Object, e As EventArgs) Handles btnModify_車籍.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法修改")
            Exit Sub
        End If
        btnModify_Click(sender, e)
    End Sub

    '刪除-系統設定-權限設定
    Private Sub btnDel_權限_Click(sender As Object, e As EventArgs) Handles btnDel_權限.Click
        If txtName_權限.Text = "" Then
            MsgBox("請選擇對象")
            Exit Sub
        End If
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteTable("密碼資料表", $"{txtName_權限.Tag} = '{txtName_權限.Text}'") Then
            btnClear_權限.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    'dgv點擊-廠商資料、客戶資料、車籍資料
    Private Sub dgv_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv廠商.CellMouseClick, dgv客戶.CellMouseClick, dgv車籍.CellMouseClick
        Dim dgv As DataGridView = sender
        If dgv.SelectedRows.Count < 1 Then Exit Sub
        Dim tp = dgv.Parent
        Dim selectRow = dgv.SelectedRows(0)
        GetDataToControls(tp, selectRow)
    End Sub

    'dgv點擊-貨品資料
    Private Sub dgv貨品_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv貨品.CellMouseClick
        Dim dgv As DataGridView = sender
        If dgv.SelectedRows.Count < 1 Then Exit Sub
        Dim row = dgv貨品.SelectedRows(0)
        GetDataToControls(tp貨品, row)
        If row.Cells("噸數設定").Value Then
            rdoTSet.Checked = True
        Else
            rdoTUnset.Checked = True
        End If
    End Sub

    '查詢-過磅作業
    Private Sub btnQuery_過磅_Click(sender As Object, e As EventArgs) Handles btnQuery_過磅.Click
        Call frmScaleQuery.Show()
    End Sub

    ''' <summary>
    ''' 檢查Insert前的動作,並Insert(只取得TextBox的資料)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="required">必填欄位 key:名稱 value:控制項</param>
    ''' <param name="duplication">主鍵 會自己抓取ctrl.tag(欄位名稱) ctrl.text(值)</param>
    ''' <param name="table"></param>
    ''' <returns></returns>
    Private Function CheckInsert(sender As Button, required As Dictionary(Of String, Object), duplication As List(Of Object), table As String) As Boolean
        Dim tp = sender.Parent
        'If Not CheckText(tp, required.ToList) Then Return False
        If Not CheckRequiredCol(required) Then Return False
        If Not CheckDuplication(GetTableAllData(table), duplication, tp.Controls.OfType(Of DataGridView).First) Then Return False
        '取得各欄位的值
        Dim dic As New Dictionary(Of String, Object)
        tp.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) dic.Add(txt.Tag.ToString, txt.Text))
        '取得GruopBox裡的TextBox
        For Each grp In tp.Controls.OfType(Of GroupBox)
            grpDecimal.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) dic.Add(txt.Tag.ToString, txt.Text))
        Next
        If Not InserTable(table, dic) Then
            Return False
        End If
        Return True
    End Function

    '修改-廠商資料,客戶資料
    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify_廠商.Click, btnModify_客戶.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法修改")
            Exit Sub
        End If
        Dim tp = CType(sender, Button).Parent
        Dim table As String = ""
        Dim btn As Button = Nothing
        Dim condition() As String = {}
        Dim required = New List(Of String)
        Select Case tp.Text
            Case "廠商資料"
                table = "廠商資料表"
                btn = btnClear_廠商
                condition = {"代號", txtNo_廠商.Text}
                required.AddRange({"全銜", "代號", "簡稱"})
            Case "客戶資料"
                table = "客戶資料表"
                btn = btnClear_客戶
                condition = {"代號", txtNo_客戶.Text}
                required.AddRange({"全銜", "代號", "簡稱"})
            Case "車籍資料"
                table = "車籍資料表"
                btn = btnClear_車籍
                condition = {"車號", txtNo_車籍.Text}
                required.AddRange({"車號", "車主"})
        End Select

        If Not CheckText(tp, required) Then Exit Sub
        '取得欄位資料
        Dim dic As New Dictionary(Of String, String)
        tp.Controls.OfType(Of TextBox).Where(Function(txt) txt.Tag <> condition(0)).ToList.ForEach(Sub(txt) dic.Add(txt.Tag.ToString, txt.Text))

        If Not UpdateTable(table, dic, $"{condition(0)} = '{condition(1)}'") Then Exit Sub

        btn.PerformClick()
        MsgBox("修改成功")
    End Sub

    '刪除-過磅作業-二次過磅
    Private Sub btnDel_二次過磅_Click(sender As Object, e As EventArgs) Handles btnDel_二次過磅.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法刪除")
            Exit Sub
        End If
        If dgv二次過磅.SelectedRows.Count = 0 Then Exit Sub
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        Dim row = dgv二次過磅.SelectedRows(0)
        Dim dic As New Dictionary(Of String, String) From {{"車牌號碼", row.Cells("車牌號碼").Value.ToString}}
        If DeleteTable("二次過磅暫存資料表", dic) Then
            btnClear_過磅.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '刪除-過磅作業
    Private Sub btnDel_過磅_Click(sender As Object, e As EventArgs) Handles btnDel_過磅.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法刪除")
            Exit Sub
        End If
        If String.IsNullOrEmpty(txtRcepNo.Text) Then
            MsgBox("請選擇對象")
            Exit Sub
        End If
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        Dim dic As New Dictionary(Of String, String) From {{"磅單序號", txtRcepNo.Text}}
        If DeleteTable("過磅資料表", dic) Then
            btnClear_過磅.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '刪除-廠商資料,客戶資料,車籍資料,貨品資料,廠商資料-專案,客戶資料-工程
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel_廠商.Click, btnDel_客戶.Click, btnDel_車籍.Click, btnDel_貨品.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法刪除")
            Exit Sub
        End If
        Dim ctrl = sender.Parent
        Dim table As String = ""
        Dim btn As Button = Nothing
        Dim condition() As String = {}
        Select Case ctrl.Text
            Case "廠商資料"
                table = "廠商資料表"
                btn = btnClear_廠商
                condition = {"代號", txtNo_廠商.Text}
            Case "客戶資料"
                table = "客戶資料表"
                btn = btnClear_客戶
                condition = {"代號", txtNo_客戶.Text}
            Case "車籍資料"
                table = "車籍資料表"
                btn = btnClear_車籍
                condition = {"車號", txtNo_車籍.Text}
            Case "貨品資料"
                table = "產品資料表"
                btn = btnClear_貨品
                condition = {"代號", txtNo_貨品.Text}
        End Select

        If String.IsNullOrWhiteSpace(condition(1)) Then
            MsgBox("請選擇刪除對象")
            Exit Sub
        End If

        If MsgBox($"你確定要刪除 {condition(1)} ?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        If Not DeleteTable(table, $"{condition(0)} = '{condition(1)}'") Then Exit Sub

        btn.PerformClick()
        MsgBox("刪除成功")
    End Sub

    '清除-廠商資料,客戶資料
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear_廠商.Click, btnClear_客戶.Click
        Dim tp = sender.Parent
        ClearControl(tp)
        BtnClear(tp)
    End Sub

    '清除-貨品資料
    Private Sub btnClear_貨品_Click(sender As Object, e As EventArgs) Handles btnClear_貨品.Click
        Dim tp = CType(sender, Button).Parent
        tp.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Clear())
        For Each grp In tp.Controls.OfType(Of GroupBox)
            grpDecimal.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Clear())
        Next
        BtnClear(tp)
    End Sub

    Private Sub BtnClear(tp As TabPage)
        Dim table As String = ""
        Dim dgv As DataGridView = Nothing
        Select Case tp.Text
            Case "廠商資料"
                table = "廠商資料表"
                dgv = dgv廠商
            Case "客戶資料"
                table = "客戶資料表"
                dgv = dgv客戶
            Case "車籍資料"
                table = "車籍資料表"
                dgv = dgv車籍
            Case "貨品資料"
                table = "產品資料表"
                dgv = dgv貨品
        End Select
        dgv.DataSource = SelectTable(GetTableAllData(table))
    End Sub

    '清除-車籍資料
    Private Sub btnClear_車籍_Click(sender As Object, e As EventArgs) Handles btnClear_車籍.Click
        btnClear_Click(sender, e)
        Init車籍()
    End Sub

    '查詢-廠商資料,客戶資料,車籍資料
    Private Sub btnQuery_廠商_Click(sender As Object, e As EventArgs) Handles btnQuery_廠商.Click, btnQuery_客戶.Click, btnQuery_車籍.Click, btnQuery_貨品.Click
        Dim lst As New List(Of String)
        Dim dgv As DataGridView = Nothing
        Dim table As String = ""
        Select Case sender.Parent.Text
            Case "廠商資料"
                If Not String.IsNullOrWhiteSpace(txtName_廠商.Text) Then lst.Add($"全銜 LIKE '%{txtName_廠商.Text}%'")
                If Not String.IsNullOrWhiteSpace(txtNo_廠商.Text) Then lst.Add($"代號 LIKE '%{txtNo_廠商.Text}%'")
                If Not String.IsNullOrWhiteSpace(txtAka_廠商.Text) Then lst.Add($"簡稱 LIKE '%{txtAka_廠商.Text}%'")
                dgv = dgv廠商
                table = "廠商資料表"
            Case "客戶資料"
                If Not String.IsNullOrWhiteSpace(txtName_客戶.Text) Then lst.Add($"全銜 LIKE '%{txtName_客戶.Text}%'")
                If Not String.IsNullOrWhiteSpace(txtNo_客戶.Text) Then lst.Add($"代號 LIKE '%{txtNo_客戶.Text}%'")
                If Not String.IsNullOrWhiteSpace(txtAka_客戶.Text) Then lst.Add($"簡稱 LIKE '%{txtAka_客戶.Text}%'")
                dgv = dgv客戶
                table = "客戶資料表"
            Case "車籍資料"
                If Not String.IsNullOrWhiteSpace(txtNo_車籍.Text) Then lst.Add($"車號 LIKE '%{txtNo_車籍.Text}%'")
                If Not String.IsNullOrWhiteSpace(txt車主.Text) Then lst.Add($"車主 LIKE '%{txt車主.Text}%'")
                dgv = dgv車籍
                table = "車籍資料表"
            Case "貨品資料"
                If Not String.IsNullOrWhiteSpace(txtNo_貨品.Text) Then lst.Add($"代號 LIKE '%{txtNo_貨品.Text}%'")
                If Not String.IsNullOrWhiteSpace(txtName_貨品.Text) Then lst.Add($"品名 LIKE '%{txtName_貨品.Text}%'")
                dgv = dgv貨品
                table = "產品資料表"
        End Select
        If lst.Count = 0 Then Exit Sub
        dgv.DataSource = SelectTable(GetTableAllData(table) + $" WHERE " + String.Join(" AND ", lst))
    End Sub

    Sub ReadDataGridWidth(dgv As String)
        Dim myObject As Object

        myObject = Me.Controls.Find(dgv, True)

        Dim newDGV = CType(myObject(0), DataGridView)
        With newDGV
            Dim tmpWidth As String
            Dim objStreamReader As StreamReader
            Try
                objStreamReader = New StreamReader(dgv + ".set", False)
                tmpWidth = objStreamReader.ReadLine()
                objStreamReader.Close()
                Dim tmpW() = Split(tmpWidth, ",")
                For j = 1 To .ColumnCount
                    .Columns(j - 1).Width = tmpW(j - 1)
                Next
            Catch ex As Exception

            End Try
        End With
    End Sub

    Sub SaveDataGridWidth(dgv As String)
        Dim myObject As Object

        myObject = Me.Controls.Find(dgv, True)

        Dim newDGV = CType(myObject(0), DataGridView)
        With newDGV
            Dim tmpWidth As String
            tmpWidth = .Columns(0).Width.ToString
            For j = 2 To .ColumnCount
                tmpWidth = tmpWidth + "," + .Columns(j - 1).Width.ToString
            Next
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(dgv + ".set", False)
            objStreamWriter.WriteLine(tmpWidth)
            objStreamWriter.Close()
        End With
    End Sub

    Private Sub dgv廠商_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dgv廠商.ColumnWidthChanged
        SaveDataGridWidth(sender.name)
    End Sub

    Private Sub lst廠商_Click(sender As Object, e As EventArgs) Handles lst廠商.Click, lst客戶.Click
        txt車主.Clear()
        Dim lst = CType(sender, ListBox)
        txt車主.Text = lst.GetItemText(lst.SelectedItem)
    End Sub

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        portClose = True
        End
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabMain.SelectedIndexChanged
        If tabMain.SelectedTab.Name = "tpLogout" Then
            If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then
                Hide()
                portClose = True
                LoginForm1.Show()
            End If
        End If
    End Sub

    ''過磅作業-空重輸入時間
    'Private Sub btnEmptyCarTime_Click(sender As Object, e As EventArgs) Handles btnEmptyCarTime.Click
    '    txtLoudTime_Empty.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
    'End Sub

    ''過磅作業-總重輸入時間
    'Private Sub btnFullCarTime_Click(sender As Object, e As EventArgs) Handles btnFullCarTime.Click
    '    txtLoudTime_Total.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
    'End Sub

    '過磅作業-空重輸入時間
    Private Sub txtEmptyCar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEmptyCar.KeyPress
        If txtEmptyCar.ReadOnly = True Then Exit Sub
        txtLoudTime_Empty.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
        If String.IsNullOrWhiteSpace(txtEmptyCar.Text) Then txtLoudTime_Empty.Clear()
    End Sub

    '過磅作業-總重輸入時間
    Private Sub txtTotalWeight_KeyUp(sender As Object, e As KeyEventArgs) Handles txtTotalWeight.KeyUp
        '手動模式觸發
        If txtTotalWeight.ReadOnly = True Then Exit Sub
        txtLoudTime_Total.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
        If String.IsNullOrWhiteSpace(txtTotalWeight.Text) Then txtLoudTime_Total.Clear()
    End Sub
End Class
