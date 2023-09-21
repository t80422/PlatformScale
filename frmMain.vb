Imports System.IO
Imports System.IO.Ports
Imports System.Text.RegularExpressions
Imports iText.Kernel.Pdf
Imports iText.Html2pdf
Imports iText.Layout
Imports iText.Html2pdf.Resolver.Font
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms.Application
Imports System.Runtime.InteropServices

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
        '檢查資料夾是否存在
        Dim folderPath As String = Path.Combine(StartupPath, "Report")
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

        '校正dtp時間
        dtp過磅.Value = Now

        InitRcepStyle()
    End Sub

    ''' <summary>
    ''' 初始化 系統設定-過磅單樣式
    ''' </summary>
    Private Sub InitRcepStyle()
        '設定 系統設定-過磅單樣式 cmb
        Dim dic = New Dictionary(Of String, String) From {
            {"直式", "A"},
            {"橫式", "B"}
        }
        With cmbRcepStyle
            For Each kvp In dic
                .Items.Add(kvp)
            Next
            .DisplayMember = "Key"
        End With

        '載入設定檔
        Dim filePath = Path.Combine(StartupPath, "RcrpStyle.set")
        If Not File.Exists(filePath) Then
            File.Create(filePath).Close()
            Exit Sub
        Else
            For Each kvp In dic
                If kvp.Value = File.ReadAllText(filePath) Then
                    cmbRcepStyle.SelectedItem = kvp
                    Exit For
                End If
            Next
        End If
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
        If Not IsDBNull(row("埠口1")) AndAlso row("埠口1") <> 0 Then
            serialPortA = New SerialPort("COM" + row("埠口1").ToString, 9600, Parity.None, 7, StopBits.One)
            AddHandler serialPortA.DataReceived, AddressOf SerialPortA_DataReceived
            Try
                If serialPortA.IsOpen Then serialPortA.Close()
                serialPortA.Open()
                portClose = False
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        If Not IsDBNull(row("埠口2")) AndAlso row("埠口2") <> 0 Then
            serialPortB = New SerialPort("COM" + row("埠口2").ToString, 9600, Parity.None, 7, StopBits.One)
            AddHandler serialPortB.DataReceived, AddressOf SerialPortB_DataReceived
            Try
                If serialPortB.IsOpen Then serialPortB.Close()
                serialPortB.Open()
                portClose = False
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

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

    Private Sub cmbCarNo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCarNo.SelectedIndexChanged
        If cmbCarNo.SelectedIndex = -1 Then txtEmptyCar.Clear()
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
        Dim d = Now
        Dim dt = SelectTable($"SELECT 磅單序號 FROM 過磅資料表 WHERE 過磅日期 = '{d:yyyy/MM/dd}' " &
                                                "UNION " &
                                              $"SELECT 磅單序號 FROM 二次過磅暫存資料表 WHERE 過磅日期 = '{d:yyyy/MM/dd}' " &
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
        If dgv過磅.SelectedRows.Count = 0 Then
            MsgBox("無磅單可列印")
            Exit Sub
        End If

        Dim id As String = dgv過磅.SelectedRows(0).Cells("磅單序號").Value
        PrintRcep(id)
    End Sub

    ''' <summary>
    ''' 列印過磅單
    ''' </summary>
    ''' <param name="id">磅單序號</param>
    Private Sub PrintRcep(id As String)
        Cursor = Cursors.WaitCursor

        If cmbRcepStyle.SelectedIndex = -1 Then
            MsgBox("請到系統設定先設定過磅單樣式")
            GoTo Finish
        End If

        Dim type = cmbRcepStyle.SelectedItem.value
        Dim data = SelectTable($"SELECT * FROM 過磅資料表 a LEFT JOIN 車籍資料表 b ON a.車牌號碼 = b.車號 WHERE a.磅單序號 = '{id}'")
        Dim fileName As String = ""

        Select Case type
            Case "A"
                fileName = "直式.html"
            Case "B"
                fileName = "橫式.html"
        End Select

        Dim folder = Path.Combine(StartupPath, "Rcep")
        Dim filePath = Path.Combine(folder, fileName)
        Dim pdfFilePath = Path.Combine(folder, "test.pdf")

        '檢查PDF是否開啟,有就關閉
        Dim processes = Process.GetProcessesByName("AcroRd32")

        For Each process In processes
            If pdfFilePath = process.MainModule.FileName Then
                process.CloseMainWindow()
                process.WaitForExit()
            End If
        Next

        Using fs = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Using sr = New StreamReader(fs)
                Dim lines = sr.ReadToEnd

                '取代文字
                Dim regex As New Regex("\[\$(.*?)\]")
                Dim matches = regex.Matches(lines)
                For Each match As Match In matches
                    Dim columnName = match.Groups(1).Value
                    Dim value As String = GetColumnValue(data, columnName)
                    lines = lines.Replace(match.Value, value)
                Next

                '另存成PDF
                Using pdf = New PdfDocument(New PdfWriter(pdfFilePath))
                    '設定中文字型
                    Dim fontProvider = New DefaultFontProvider
                    fontProvider.AddFont("c:/windows/Fonts/MSMINCHO.TTF")
                    fontProvider.AddFont("c:/windows/Fonts/STSONG.TTF")
                    Dim cp = New ConverterProperties
                    cp.SetFontProvider(fontProvider)

                    HtmlConverter.ConvertToPdf(lines, pdf, cp)
                End Using
            End Using
        End Using

        Process.Start(pdfFilePath)
Finish:
        Cursor = Cursors.Default
    End Sub

    Private Function GetColumnValue(data As DataTable, columnName As String) As String
        If Not data.Columns.Contains(columnName) Then
            Dim inout As String = data.Rows(0)("進/出")
            If columnName = "入廠時間" Then
                columnName = If(inout = "進貨", "總重載入時間", "空重載入時間")
            ElseIf columnName = "出廠時間" Then
                columnName = If(inout = "進貨", "空重載入時間", "總重載入時間")
            End If
        End If

        If IsDBNull(data.Rows(0)(columnName)) Then
            Return ""
        Else
            Return data.Rows(0)(columnName).ToString()
        End If

    End Function

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

    '廠商資料-刪除
    Private Sub btnDel_廠商_Click(sender As Object, e As EventArgs) Handles btnDel_廠商.Click
        CommonDelete(txtNo_廠商, "廠商資料表")
    End Sub

    '客戶資料-刪除
    Private Sub btnDel_客戶_Click(sender As Object, e As EventArgs) Handles btnDelete_客戶.Click
        CommonDelete(txtNo_客戶, "客戶資料表")
    End Sub

    ''' <summary>
    ''' 刪除鍵共同模組
    ''' </summary>
    ''' <param name="txt">編號的TextBox</param>
    ''' <param name="table">對應的資料表</param>
    Private Sub CommonDelete(txt As TextBox, table As String)
        If permissions = 1 Then
            MsgBox("權限不足,無法刪除")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(txt.Text) Then
            MsgBox("請選擇刪除對象")
            Exit Sub
        End If

        If MsgBox($"確定要刪除?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub

        If DeleteTable(table, $"{txt.Tag} = '{txt.Text}'") Then
            txt.Parent.Controls.OfType(Of Button).First(Function(btn) btn.Name.Contains("btnClear")).PerformClick()
            btnClear_車籍_Click(btnClear_車籍, EventArgs.Empty)
            MsgBox("刪除成功")
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
    Private Sub btnClear_過磅_Click(sender As Object, e As EventArgs) Handles btnClear_過磅.Click
        Dim btn As ButtonBase = sender
        ClearControl(btn.Parent.Controls.OfType(Of Control).Where(Function(ctrl) TypeOf ctrl IsNot GroupBox))
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
        grpInOut.Controls.OfType(Of RadioButton).Where(Function(rdo) rdo.Text = selectRow.Cells("進/出").Value).ToList.ForEach(Sub(x) x.Checked = True)
        cmbCliManu.SelectedIndex = cmbCliManu.FindStringExact(selectRow.Cells("客戶/廠商").Value)
        Dim carNo = GetCellData(selectRow, "車牌號碼")
        cmbCarNo.SelectedIndex = cmbCarNo.FindStringExact(carNo)
        cmbProduct.SelectedIndex = cmbProduct.FindStringExact(selectRow.Cells("產品名稱").Value)
        tp過磅.Controls.OfType(Of TextBox).Where(Function(txt) txt.Tag IsNot Nothing AndAlso Not IsDBNull(selectRow.Cells(txt.Tag.ToString).Value)).
            ToList.ForEach(Sub(txt) txt.Text = selectRow.Cells(txt.Tag.ToString).Value)
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
        'txtEmptyCar.Clear()
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

        '如果有空重時間與總重時間表示完成過磅
        If Not String.IsNullOrEmpty(txtLoudTime_Empty.Text) And Not String.IsNullOrEmpty(txtLoudTime_Total.Text) Then
            '檢查空重是否超過總重
            If Double.Parse(txtEmptyCar.Text) > Double.Parse(txtTotalWeight.Text) Then
                MsgBox("空重不能大於總重")
                Exit Sub
            End If

            '新增或修改
            If SelectTable($"SELECT 磅單序號 FROM 過磅資料表 WHERE {txtRcepNo.Tag} = '{txtRcepNo.Text}'").Rows.Count = 0 Then
                If Save過磅Data("過磅資料表", "insert") Then
                    Dim dic As New Dictionary(Of String, String) From {{"車牌號碼", cmbCarNo.Text}}
                    DeleteTable("二次過磅暫存資料表", dic)
                    If MsgBox("是否列印過磅單", MessageBoxButtons.YesNo) = MsgBoxResult.Yes Then PrintRcep(txtRcepNo.Text)
                Else
                    Exit Sub
                End If

            Else
                If MsgBox($"{txtRcepNo.Tag}:{txtRcepNo.Text} 已修改,是否覆蓋", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If Save過磅Data("過磅資料表", "update") Then
                        Dim dic As New Dictionary(Of String, String) From {{"車牌號碼", cmbCarNo.Text}}
                        DeleteTable("二次過磅暫存資料表", dic)
                    Else
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If

            '二次過磅(對象為未登錄空車重)
        ElseIf Not String.IsNullOrEmpty(txtLoudTime_Empty.Text) OrElse Not String.IsNullOrEmpty(txtLoudTime_Total.Text) Then
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
            If Not Save過磅Data("二次過磅暫存資料表", "insert") Then Exit Sub
        Else
            MsgBox("請先檢查重量貨載入時間")
            Exit Sub
        End If

        '臨時客戶/廠商,新增到資料表
        Dim cm As enumWho
        If cmbCliManu.SelectedIndex = -1 Then
            Dim table As String = ""

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

                '對應臨時客戶/廠商新增時要刷新
                btnClear_Click(btnClear_客戶, e)
                btnClear_Click(btnClear_廠商, e)
            End If
        End If

        '臨時車號,新增到資料表
        If cmbCarNo.SelectedIndex = -1 Then
            '檢查是否重複
            If SelectTable($"SELECT 車號 FROM 車籍資料表 WHERE 車號 = '{cmbCarNo.Text}' AND 車主 = '{cmbCliManu.Text}'").Rows.Count = 0 Then
                Dim dic As New Dictionary(Of String, Object) From {
                    {"車主", cmbCliManu.Text},
                    {"車號", cmbCarNo.Text}
                }
                InserTable("車籍資料表", dic)
            End If

            '對應臨時車號新增時要刷新
            btnClear_車籍_Click(btnClear_車籍, EventArgs.Empty)
        Else
            GoTo Finish
        End If
        SetCmbCliManu(cm)
Finish:
        btnClear_過磅.PerformClick()
        MsgBox("儲存成功")
    End Sub

    ''' <summary>
    ''' 儲存過磅資料
    ''' </summary>
    ''' <param name="table">資料表</param>
    ''' <param name="status">insert、update</param>
    ''' <returns></returns>
    Private Function Save過磅Data(table As String, status As String) As Boolean
        Dim dicRequired As New Dictionary(Of String, Object) From {
            {"廠商/客戶", cmbCliManu},
            {"產品名稱", cmbProduct},
            {"車號", cmbCarNo}
        }
        If Not CheckRequiredCol(dicRequired) Then Return False

        Dim dicData As New Dictionary(Of String, String)
        For Each ctrl In tp過磅.Controls.OfType(Of Control).Where(Function(ctrls) ctrls.Tag IsNot Nothing AndAlso ctrls.Text <> "")
            Dim ctrlType = ctrl.GetType.Name
            Dim ctrlTag = ctrl.Tag
            Dim ctrlText = ctrl.Text

            Select Case ctrlType
                Case "TextBox", "ComboBox"
                    dicData.Add(ctrlTag, ctrlText)

                Case "DateTimePicker"
                    If status = "insert" Then
                        dicData.Add(ctrlTag, DirectCast(ctrl, DateTimePicker).Value.ToString("yyyy/MM/dd"))
                    End If

                Case "Label"
                    If status = "insert" Then dicData.Add(ctrlTag, ctrlText)
            End Select
        Next

        Dim inOutRadioButton = grpInOut.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked)
        dicData.Add(grpInOut.Tag, inOutRadioButton.Text)

        Dim success As Boolean = False
        If status = "insert" Then
            success = InserTable(table, dicData)
        Else
            success = UpdateTable(table, dicData, $"磅單序號 = '{txtRcepNo.Text}'")
        End If

        Return success
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
        '輸入防呆
        If Not String.IsNullOrWhiteSpace(txtPortA.Text) AndAlso Not CheckPositiveNumber(txtPortA) Then
            MsgBox("請輸入正確的Port")
            txtPortA.Focus()
            Exit Sub
        End If
        If Not String.IsNullOrWhiteSpace(txtPortB.Text) AndAlso Not CheckPositiveNumber(txtPortB) Then
            MsgBox("請輸入正確的Port")
            txtPortB.Focus()
            Exit Sub
        End If

        '檢查是否兩個Port設定一樣
        Dim row = SelectTable("SELECT * FROM 通訊埠口資料表").Rows(0)
        If Not String.IsNullOrWhiteSpace(txtPortA.Text) Then
            If Not IsDBNull(row("埠口2")) AndAlso txtPortA.Text = row("埠口2") Then
                MsgBox("不能與 埠口2 設定一樣")
                Exit Sub
            End If
        End If
        If Not String.IsNullOrWhiteSpace(txtPortB.Text) Then
            If Not IsDBNull(row("埠口1")) AndAlso txtPortB.Text = row("埠口1") Then
                MsgBox("不能與 埠口1 設定一樣")
                Exit Sub
            End If
        End If

        Dim dic As New Dictionary(Of String, String) From {
            {txtPortA.Tag, If(String.IsNullOrWhiteSpace(txtPortA.Text), 0, txtPortA.Text)},
            {txtPortB.Tag, If(String.IsNullOrWhiteSpace(txtPortB.Text), 0, txtPortB.Text)}
        }
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
            File.Copy(StartupPath + "\db4UGWS.mdb", dbPath, True)
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
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel_車籍.Click, btnDel_貨品.Click
        If permissions = 1 Then
            MsgBox("權限不足,無法刪除")
            Exit Sub
        End If
        Dim ctrl = sender.Parent
        Dim table As String = ""
        Dim btn As Button = Nothing
        Dim condition() As String = {}
        Select Case ctrl.Text
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
            Case "貨品資料"
                table = "產品資料表"
                dgv = dgv貨品
        End Select
        dgv.DataSource = SelectTable(GetTableAllData(table))
    End Sub

    '清除-車籍資料
    Private Sub btnClear_車籍_Click(sender As Object, e As EventArgs) Handles btnClear_車籍.Click
        Dim btn As Button = sender
        ClearControl(btn.Parent)
        dgv車籍.DataSource = SelectTable(GetTableAllData("車籍資料表"))
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

    '過磅作業-總重輸入時間
    Private Sub txtTotalWeight_KeyUp(sender As Object, e As KeyEventArgs) Handles txtTotalWeight.KeyUp
        '手動模式觸發
        If txtTotalWeight.ReadOnly = True Then Exit Sub
        txtLoudTime_Total.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
        If String.IsNullOrWhiteSpace(txtTotalWeight.Text) Then txtLoudTime_Total.Clear()
    End Sub

    '過磅作業-空重輸入時間
    Private Sub txtEmptyCar_KeyUp(sender As Object, e As KeyEventArgs) Handles txtEmptyCar.KeyUp
        If txtEmptyCar.ReadOnly = True Then Exit Sub
        txtLoudTime_Empty.Text = Date.Parse(lblTime.Text).ToString("HH:mm")
        If String.IsNullOrWhiteSpace(txtEmptyCar.Text) Then txtLoudTime_Empty.Clear()
    End Sub

    '系統設定-Poty設定-dgv點擊
    Private Sub dgvPort_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvPort.CellMouseClick
        Dim row = dgvPort.SelectedRows(0)
        If row.Index = -1 Then Exit Sub
        GetDataToControls(grpPort, row)
    End Sub

    '系統設定-過磅單樣式
    Private Sub cmbRcepStyle_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbRcepStyle.SelectionChangeCommitted
        Dim filePath = Path.Combine(StartupPath, "RcrpStyle.set")
        If File.Exists(filePath) Then
            Dim item = cmbRcepStyle.SelectedItem
            File.WriteAllText(filePath, item.value)
        End If
    End Sub

    '報表-列印
    Private Sub btnPrint_report_Click(sender As Object, e As EventArgs) Handles btnPrint_report.Click
        btnPrint_report.Enabled = False
        Cursor = Cursors.WaitCursor

        Dim inOut = grpInOut_report.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text
        Dim type = grpType_report.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text
        Dim dStart As String = ""
        Dim dEnd As String = ""
        Dim dt As Data.DataTable
        Dim exl As New Application
        Dim wb As Workbook = Nothing

        Try
            '將範本檔的Sheet複製到新的Excel檔
            Dim orgWb As Workbook = exl.Workbooks.Open("C:\Users\t8042\Desktop\報表.xlsx")
            Dim orgWs As Worksheet = orgWb.Worksheets(type)
            wb = exl.Workbooks.Add
            orgWs.Copy(wb.Sheets(1))
            orgWb.Close(False)
            wb.Sheets(2).Delete

            Dim ws As Worksheet = wb.Worksheets(type)
            Dim cells As Range = ws.Cells
            Dim rowIndex As Integer = 0

            Select Case type
                Case "年度對帳單"
                    '撈資料
                    dStart = dtpStart.Value.Year.ToString
                    dt = SelectTable(
                        "SELECT 過磅日期, [客戶/廠商], 產品名稱, 車牌號碼, 淨重, 米數, 總價 FROM 過磅資料表 " &
                       $"WHERE YEAR(CDATE(過磅日期)) = '{dStart}' " &
                       $"AND [進/出] = '{inOut}'"
                                    )

                    '寫入Excel
                    '標題
                    cells(1, 1) = $"{dStart} 年度 {inOut} 對帳單"

                    '撈所有月份
                    Dim monthes = (
                        From row In dt
                        Select Date.Parse(row("過磅日期")).Month
                    ).Distinct
                    rowIndex = 3

                    For Each m In monthes
                        rowIndex = YearlyData(dt, m, cells, rowIndex)
                    Next

                Case "月對帳單"
                    '撈資料
                    Dim year As Integer = dtpStart.Value.Year
                    Dim month As Integer = dtpStart.Value.Month
                    dt = SelectTable(
                        "SELECT 過磅日期, [客戶/廠商], 磅單序號, 空重, 總重, 單價, 每米噸數, 淨重, 米數, 總價, 產品名稱, 車牌號碼 FROM 過磅資料表 " &
                        $"WHERE YEAR(CDATE(過磅日期)) = '{year}' " &
                        $"AND MONTH(CDATE(過磅日期)) = '{month}' " &
                        $"AND [進/出] = '{inOut}'"
                    )

                    '寫入Excel
                    '標題
                    cells(1, 1) = $"{year} 年 {month} 月 {inOut} 對帳單"

                    ' 撈取當月客戶購買的產品所使用的車的資料
                    Dim datas = dt.AsEnumerable().
                        GroupBy(Function(row) New With {
                            .Customer = row("客戶/廠商"),
                            .Product = row("產品名稱"),
                            .LicensePlate = row("車牌號碼")
                        }).
                        Select(Function(group) New With {
                            group.Key.Customer,
                            group.Key.Product,
                            group.Key.LicensePlate,
                            .Data = group.Select(Function(row) New With {
                                .Date = row("過磅日期"),
                                .ID = row("磅單序號"),
                                .Empty = Math.Round(row("空重"), 3),
                                .Weight = Math.Round(row("總重"), 3),
                                .UnitPrice = row("單價"),
                                .TPM = row("每米噸數"),
                                .NetWeight = Math.Round(row("淨重"), 3),
                                .Meter = Math.Round(row("米數"), 3),
                                .Price = Math.Round(row("總價"), 3)
                            })
                        })

                    rowIndex = 3

                    For Each item In datas
                        ' 客戶、產品、車號
                        cells(rowIndex, 1) = item.Customer
                        cells(rowIndex, 2) = item.Product
                        cells(rowIndex, 3) = item.LicensePlate
                        BottomLine_Cell(cells(rowIndex, 1))
                        BottomLine_Cell(cells(rowIndex, 2))
                        BottomLine_Cell(cells(rowIndex, 3))
                        rowIndex += 1

                        For Each d In item.Data
                            cells(rowIndex, 1) = d.Date
                            cells(rowIndex, 2) = d.ID
                            cells(rowIndex, 3) = d.Empty
                            cells(rowIndex, 4) = d.Weight
                            cells(rowIndex, 5) = d.UnitPrice
                            cells(rowIndex, 6) = d.TPM
                            cells(rowIndex, 7) = d.NetWeight
                            cells(rowIndex, 8) = d.Meter
                            cells(rowIndex, 9) = d.Price

                            rowIndex += 1
                        Next

                        '畫線
                        For i As Integer = 1 To 9
                            BottomLine_Cell(cells(rowIndex - 1, i))
                        Next

                        ' 小計
                        cells(rowIndex, 3).Value = "(小計)"
                        cells(rowIndex, 7).Value = item.Data.Sum(Function(d) d.NetWeight)
                        cells(rowIndex, 8).Value = item.Data.Sum(Function(d) d.Meter)

                        rowIndex += 2
                    Next

                Case "日對帳單"
                    '撈資料
                    Dim dateSelect As String = dtpStart.Value.ToString("yyyy/MM/dd")

                    dt = SelectTable(
                        "SELECT 過磅日期, [客戶/廠商], 磅單序號, 空重, 總重, 單價, 每米噸數, 淨重, 米數, 總價, 產品名稱, 車牌號碼 FROM 過磅資料表 " &
                        $"WHERE 過磅日期 = '{dateSelect}' " &
                        $"AND [進/出] = '{inOut}'"
                    )

                    '寫入Excel
                    '標題
                    cells(1, 1) = $"{dtpStart.Value.Year} 年 {dtpStart.Value.Month} 月 {dtpStart.Value.Day} 日 {inOut} 對帳單"

                    ' 撈取當月客戶購買的產品所使用的車的資料
                    Dim datas = dt.AsEnumerable().
                        GroupBy(Function(row) New With {
                            .Customer = row("客戶/廠商"),
                            .Product = row("產品名稱"),
                            .LicensePlate = row("車牌號碼")
                        }).
                        Select(Function(group) New With {
                            group.Key.Customer,
                            group.Key.Product,
                            group.Key.LicensePlate,
                            .Data = group.Select(Function(row) New With {
                                .Date = row("過磅日期"),
                                .ID = row("磅單序號"),
                                .Empty = Math.Round(row("空重"), 3),
                                .Weight = Math.Round(row("總重"), 3),
                                .UnitPrice = row("單價"),
                                .TPM = row("每米噸數"),
                                .NetWeight = Math.Round(row("淨重"), 3),
                                .Meter = Math.Round(row("米數"), 3),
                                .Price = Math.Round(row("總價"), 3)
                            })
                        })

                    rowIndex = 3

                    For Each item In datas
                        ' 客戶、產品、車號
                        cells(rowIndex, 1) = item.Customer
                        cells(rowIndex, 2) = item.Product
                        cells(rowIndex, 3) = item.LicensePlate
                        BottomLine_Cell(cells(rowIndex, 1))
                        BottomLine_Cell(cells(rowIndex, 2))
                        BottomLine_Cell(cells(rowIndex, 3))
                        rowIndex += 1

                        For Each d In item.Data
                            cells(rowIndex, 1) = d.Date
                            cells(rowIndex, 2) = d.ID
                            cells(rowIndex, 3) = d.Empty
                            cells(rowIndex, 4) = d.Weight
                            cells(rowIndex, 5) = d.UnitPrice
                            cells(rowIndex, 6) = d.TPM
                            cells(rowIndex, 7) = d.NetWeight
                            cells(rowIndex, 8) = d.Meter
                            cells(rowIndex, 9) = d.Price

                            rowIndex += 1
                        Next

                        '畫線
                        For i As Integer = 1 To 9
                            BottomLine_Cell(cells(rowIndex - 1, i))
                        Next

                        ' 小計
                        cells(rowIndex, 3).Value = "(小計)"
                        cells(rowIndex, 7).Value = item.Data.Sum(Function(d) d.NetWeight)
                        cells(rowIndex, 8).Value = item.Data.Sum(Function(d) d.Meter)

                        rowIndex += 2
                    Next

                Case Else
                    MsgBox("無效的選項", MsgBoxStyle.Exclamation, Title:="報表-列印")
                    Exit Sub
            End Select

            ws.SaveAs("C:\Users\t8042\Desktop\test.xlsx")
            Marshal.ReleaseComObject(ws)

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            If wb IsNot Nothing Then
                wb.Close(False)
                Marshal.ReleaseComObject(wb)
            End If

            exl.Quit()
            Marshal.ReleaseComObject(exl)

            btnPrint_report.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 將儲存格加上細的下框線
    ''' </summary>
    ''' <param name="cell">目標儲存格</param>
    Private Sub BottomLine_Cell(cell As Range)
        cell.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        cell.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
    End Sub

    ''' <summary>
    ''' 將儲存格加上細的上框線
    ''' </summary>
    ''' <param name="cell"></param>
    Private Sub TopLine_Cell(cell As Range)
        cell.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        cell.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
    End Sub

    Private Function YearlyData(dt As Data.DataTable, month As Integer, cells As Range, startRowIndex As Integer) As Integer
        ' 月份
        cells(startRowIndex, 1) = month & " 月"
        BottomLine_Cell(cells(startRowIndex, 1))

        Dim rowIndex = startRowIndex + 1

        ' 撈取當月客戶購買的產品所使用的車的資料
        Dim datas = dt.AsEnumerable().
            Where(Function(row) Date.Parse(row("過磅日期")).Month = month).
            GroupBy(Function(row) New With {
                .Customer = row("客戶/廠商"),
                .Product = row("產品名稱"),
                .LicensePlate = row("車牌號碼")
            }).
            Select(Function(group) New With {
                group.Key.Customer,
                group.Key.Product,
                group.Key.LicensePlate,
                .Data = group.Select(Function(row) New With {
                    .Weight = Math.Round(row("淨重"), 3),
                    .Meters = Math.Round(row("米數"), 3),
                    .Price = Math.Round(row("總價"), 3)
                })
            })

        For Each item In datas
            ' 客戶、產品、車號
            cells(rowIndex, 1) = item.Customer
            cells(rowIndex, 2) = item.Product
            cells(rowIndex, 3) = item.LicensePlate
            BottomLine_Cell(cells(rowIndex, 1))
            BottomLine_Cell(cells(rowIndex, 2))
            BottomLine_Cell(cells(rowIndex, 3))
            rowIndex += 1

            For Each d In item.Data
                ' 資料
                Dim wt = d.Weight
                Dim mt = d.Meters
                Dim pc = d.Price

                cells(rowIndex, 4).Value = wt
                cells(rowIndex, 5).Value = mt
                cells(rowIndex, 6).Value = pc
                rowIndex += 1
            Next

            ' 小計
            cells(rowIndex, 3).Value = "(小計)"
            cells(rowIndex, 4).Value = item.Data.Sum(Function(d) d.Weight)
            TopLine_Cell(cells(rowIndex, 4))
            cells(rowIndex, 5).Value = item.Data.Sum(Function(d) d.Meters)
            TopLine_Cell(cells(rowIndex, 5))
            cells(rowIndex, 6).Value = item.Data.Sum(Function(d) d.Price)
            TopLine_Cell(cells(rowIndex, 6))
            cells(rowIndex, 7).Value = item.Data.Count()
            TopLine_Cell(cells(rowIndex, 7))

            rowIndex += 2
        Next

        Return rowIndex
    End Function
End Class
