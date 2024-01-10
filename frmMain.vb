Imports System.IO
Imports System.IO.Ports
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Forms.Application
Imports iText.Html2pdf
Imports iText.Html2pdf.Resolver.Font
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf
Imports iText.Layout
Imports Microsoft.Office.Interop.Excel
Imports PlatformScale.ReportGenerators
Imports Application = System.Windows.Forms.Application
Imports DataTable = System.Data.DataTable
Imports Path = System.IO.Path
Imports TextBox = System.Windows.Forms.TextBox

Public Class frmMain
    Public permissions As Integer '權限等級
    Public user As String '使用者
    Private serialPortA As SerialPort
    Private serialPortB As SerialPort
    Private nowScale As String
    Private portClose As Boolean
    Private spAClose As Integer 'Port停止傳訊息就表示關閉了,此變數用來紀錄逾時
    Private spBClose As Integer 'Port停止傳訊息就表示關閉了,此變數用來紀錄逾時
    Private tempModify As Date '暫存最後更新時間

    Private Enum enumWho
        客戶
        廠商
    End Enum

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '檢查資料夾是否存在
        Dim folderPath As String = IO.Path.Combine(StartupPath, "Report")
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

        InitReportCombobox()

        tempModify = GetModifyTime()

        SetCheckTime()

    End Sub

    ''' <summary>
    ''' 設定資料更新頻率
    ''' </summary>
    Private Sub SetCheckTime()
        Dim sec = GetCheckTime()
        txtDBCheck.Text = sec
        tmrCheckModify.Interval = sec * 1000
    End Sub

    ''' <summary>
    ''' 取得資料更新頻率
    ''' </summary>
    ''' <returns></returns>
    Private Function GetCheckTime() As Integer
        Return SelectTable("SELECT 檢查時間 FROM 資料更新").Rows(0)("檢查時間")
    End Function

    Private Sub InitReportCombobox()
        cmbProduct_report.Items.Add("全部")
        cmbProduct_report.Items.AddRange(SelectTable("SELECT 品名 FROM 產品資料表").AsEnumerable().Select(Function(row) row("品名")).ToArray())
        cmbProduct_report.SelectedIndex = 0

        cmbCliSup_report.Items.Add("全部")
        cmbCliSup_report.Items.AddRange(SelectTable($"SELECT 簡稱 FROM 客戶資料表").AsEnumerable().Select(Function(row) row("簡稱")).ToArray())
        cmbCliSup_report.SelectedIndex = 0

        cmbCarNo_report.Items.Add("全部")
        cmbCarNo_report.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' 初始化 系統設定-過磅單樣式
    ''' </summary>
    Private Sub InitRcepStyle()
        '設定 系統設定-過磅單樣式 cmb
        Dim dic = New Dictionary(Of String, String) From {
        {"直式", "A"},
        {"橫式", "B"},
        {"直式三聯", "C"}
    }

        With cmbRcepStyle
            For Each kvp In dic
                .Items.Add(New KeyValuePair(Of String, String)(kvp.Key, kvp.Value))
            Next
            .DisplayMember = "Key"
            .ValueMember = "Value"
        End With

        '載入設定檔
        Dim filePath = IO.Path.Combine(StartupPath, "RcrpStyle.set")
        If Not File.Exists(filePath) Then
            File.Create(filePath).Close()
            Exit Sub
        Else
            Dim lines = File.ReadAllLines(filePath)
            For Each line In lines
                Dim parts = Split(line, ":")
                Select Case parts(0)
                    Case "type"
                        Dim selectedType = parts(1)
                        For Each item As KeyValuePair(Of String, String) In cmbRcepStyle.Items
                            If item.Value = selectedType Then
                                cmbRcepStyle.SelectedItem = item
                                Exit For
                            End If
                        Next
                    Case "title"
                        chkCustomizeTitle.Checked = Convert.ToBoolean(parts(1))
                    Case "text"
                        txtCustomizeTitle.Text = parts(1)
                    Case Else
                        ' 其他情況
                End Select
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
        lblA.Size = New System.Drawing.Size(50, 50)
        lblA.Region = New Region(aCircle)
        lblB.Size = New System.Drawing.Size(50, 50)
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
                    If .BackColor = System.Drawing.Color.White Then
                        .BackColor = System.Drawing.Color.Red
                    Else
                        .BackColor = System.Drawing.Color.White
                    End If
                End With
            Case "B"
                With lblB
                    If .BackColor = System.Drawing.Color.White Then
                        .BackColor = System.Drawing.Color.Red
                    Else
                        .BackColor = System.Drawing.Color.White
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
                lblB.BackColor = System.Drawing.Color.White
            Case "B"
                nowScale = AorB
                lblA.BackColor = System.Drawing.Color.White
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

    Private Sub LoadWeight(weight As TextBox, time As System.Windows.Forms.TextBox)
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
            Cursor = Cursors.Default
            Return
        End If

        Dim type = cmbRcepStyle.SelectedItem.value
        Dim data = SelectTable($"SELECT * FROM 過磅資料表 a LEFT JOIN 車籍資料表 b ON a.車牌號碼 = b.車號 WHERE a.磅單序號 = '{id}'")
        Dim fileName As String = ""

        Select Case type
            Case "A"
                fileName = "直式.html"
            Case "B"
                fileName = "橫式.html"
            Case "C"
                fileName = "直式三聯.html"
        End Select

        Dim folder = IO.Path.Combine(StartupPath, "Rcep")
        Dim filePath = IO.Path.Combine(folder, fileName)
        Dim pdfFilePath = IO.Path.Combine(folder, "test.pdf")

        CloseOpenPDF(pdfFilePath)

        Using fs = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Using sr = New StreamReader(fs)
                Dim lines = sr.ReadToEnd

                '取代文字
                lines = ReplaceTemplateText(lines, data)

                '另存成PDF
                SaveAsPDF(lines, pdfFilePath)
            End Using
        End Using

        Dim printDialog As New PrintDialog

        PrintPDF(pdfFilePath, type)

        Cursor = Cursors.Default
    End Sub

    Private Sub CloseOpenPDF(pdfFilePath As String)
        Dim processes = Process.GetProcessesByName("AcroRd32")
        For Each process In processes
            If pdfFilePath = process.MainModule.FileName Then
                process.CloseMainWindow()
                process.WaitForExit()
            End If
        Next
    End Sub

    Private Function ReplaceTemplateText(template As String, data As DataTable) As String
        Dim regex As New Regex("\[\$(.*?)\]")
        Dim matches = regex.Matches(template)

        For Each match As Match In matches
            Dim columnName = match.Groups(1).Value
            Dim replacement As String

            If columnName = "抬頭" Then
                replacement = If(chkCustomizeTitle.Checked, txtCustomizeTitle.Text, "")
            Else
                replacement = GetColumnValue(data, columnName)
            End If

            template = template.Replace(match.Value, replacement)
        Next

        Return template
    End Function

    Private Sub SaveAsPDF(htmlContent As String, pdfFilePath As String)
        Using pdf = New iText.Kernel.Pdf.PdfDocument(New PdfWriter(pdfFilePath))
            Dim fontProvider = New DefaultFontProvider(False, False, False)
            fontProvider.AddFont("c:/windows/Fonts/KAIU.TTF")
            fontProvider.AddFont("c:/windows/Fonts/msjhbd.ttf")

            Dim cp = New ConverterProperties
            cp.SetFontProvider(fontProvider)

            HtmlConverter.ConvertToPdf(htmlContent, pdf, cp)
        End Using
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

        ClearControl(btn.Parent.Controls.OfType(Of Control).Where(Function(ctrl) ctrl.GetType.Name <> "GroupBox"))

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
        lblWarningModify.Visible = False

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
        UpdateModifyTime()
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

        Dim sourcePath As String

        Using reader As New StreamReader(Path.Combine(StartupPath, "DB.set"))
            sourcePath = reader.ReadLine
        End Using

        Dim dbPath = lblRemote.Text & "\db4UGWS.mdb"
        Dim logPath = lblRemote.Text + "\DBBackUp.log"

        If Not String.IsNullOrEmpty(sourcePath) Then
            Try
                File.Copy(sourcePath, dbPath, True)
                File.AppendAllText(logPath, Now.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf)
                MsgBox("備份成功")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("未指定來源資料庫")
        End If

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
            UpdateModifyTime()
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
        Try
            End
        Catch ex As Exception

        End Try
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

    Private Sub btnPrint_report_Click(sender As Object, e As EventArgs) Handles btnPrint_report.Click
        Cursor = Cursors.WaitCursor
        btnPrint_report.Enabled = False

        Dim exlReport As New ExcelReportGenerator(StartupPath, chkExcel.Checked)

        Dim dic As New Dictionary(Of String, String) From {
            {"產品名稱", cmbProduct_report.Text},
            {"[客戶/廠商]", cmbCliSup_report.Text},
            {"車牌號碼", cmbCarNo_report.Text}
        }

        Try
            Dim inOut = grpInOut_report.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text
            Dim type = grpType_report.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text

            exlReport.CreateNewReport(type)

            Select Case type
                Case "年度對帳單"
                    exlReport.GenerateYearlyStatement(nudYear.Value, inOut, dic)

                Case "月對帳單"
                    exlReport.GenerateMonthlyStatement(nudYear.Value, nudMonth.Value, inOut, dic)

                Case "日報統計明細表"
                    exlReport.GenerateDailyStatement(nudYear.Value, nudMonth.Value, nudDay_start.Value, nudDay_end.Value, inOut, dic)

                Case "月報統計表(總量)"
                    exlReport.GenerateMonthlyReport(nudYear.Value, nudMonth.Value, nudDay_start.Value, nudDay_end.Value, inOut, dic)

                Case "月報統計表(產品)"
                    Dim dicMonthProd = New Dictionary(Of String, String) From {
                        {"[客戶/廠商]", cmbCliSup_report.Text},
                        {"車牌號碼", cmbCarNo_report.Text}
                    }

                    exlReport.GenerateMonthlyProductStats(nudYear.Value, nudMonth.Value, nudDay_start.Value, nudDay_end.Value, inOut, dicMonthProd)

                Case "日報統計表(產品)"
                    Dim startDate = New Date(nudYear.Value, nudMonth.Value, nudDay_start.Value)
                    Dim endDate = New Date(nudYear.Value, nudMonth.Value, nudDay_end.Value)

                    exlReport.GenerateDailyProductStats(startDate.ToString("yyyy/MM/dd"), endDate.ToString("yyyy/MM/dd"), inOut, dic)

                Case "日統計明細表(產品 客戶)"
                    Dim startDate = New Date(nudYear.Value, nudMonth.Value, nudDay_start.Value).ToString("yyyy/MM/dd")
                    Dim endDate = New Date(nudYear.Value, nudMonth.Value, nudDay_end.Value).ToString("yyyy/MM/dd")

                    exlReport.GenerateDailyProductCustomerStats(startDate, endDate, inOut, dic)

                Case "日報統計表(客戶 產品)"
                    Dim startDate = New Date(nudYear.Value, nudMonth.Value, nudDay_start.Value).ToString("yyyy/MM/dd")
                    Dim endDate = New Date(nudYear.Value, nudMonth.Value, nudDay_end.Value).ToString("yyyy/MM/dd")

                    exlReport.GenerateDailyCustomerProductStats(startDate, endDate, inOut, dic)

                Case "過磅單日統計表"
                    exlReport.GenerateWeighingDailyReport(nudYear.Value, nudMonth.Value, nudDay_start.Value, nudDay_end.Value, inOut, dic)

                Case Else

            End Select

            exlReport.SaveReport(type)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        exlReport.Close()
        MsgBox("完成")

        btnPrint_report.Enabled = True
        Cursor = Cursors.Default
    End Sub

    Private Sub nudYearMonth_ValueChanged(sender As Object, e As EventArgs) Handles nudMonth.ValueChanged, nudYear.ValueChanged
        If nudYear.Value = 0 OrElse nudMonth.Value = 0 Then Exit Sub

        nudDay_start.Maximum = Date.DaysInMonth(nudYear.Value, nudMonth.Value)
        nudDay_end.Maximum = Date.DaysInMonth(nudYear.Value, nudMonth.Value)
        nudDay_end.Value = nudDay_end.Maximum
    End Sub

    Private Sub CliSupButton_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCustomer.CheckedChanged, rdoSupplier.CheckedChanged
        Dim rdo As RadioButton = sender

        If String.IsNullOrEmpty(rdo.Text) OrElse Not rdo.Checked Then Exit Sub

        Dim table As String = ""

        Select Case rdo.Text
            Case "出貨"
                lblCliSup.Text = "客戶"
                table = "客戶資料表"
            Case "進貨"
                lblCliSup.Text = "廠商"
                table = "廠商資料表"
        End Select

        cmbCliSup_report.Items.Clear()
        cmbCliSup_report.Items.Add("全部")
        cmbCliSup_report.Items.AddRange(SelectTable($"SELECT 簡稱 FROM {table}").AsEnumerable().Select(Function(row) row("簡稱")).ToArray())
        cmbCliSup_report.SelectedIndex = 0
    End Sub

    Private Sub cmbCliSup_report_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbCliSup_report.SelectionChangeCommitted
        cmbCarNo_report.Items.Clear()
        cmbCarNo_report.Items.Add("全部")
        cmbCarNo_report.SelectedIndex = 0

        If cmbCliSup_report.Text <> "全部" Then
            cmbCarNo_report.Items.AddRange(SelectTable($"SELECT 車號 FROM 車籍資料表 WHERE 車主 = '{cmbCliSup_report.Text}'").AsEnumerable().Select(Function(row) row("車號")).ToArray())
        End If
    End Sub

    Private Sub btnSave_rcep_Click(sender As Object, e As EventArgs) Handles btnSave_rcep.Click
        Try
            Dim filePath = IO.Path.Combine(StartupPath, "RcrpStyle.set")
            Dim kvp As KeyValuePair(Of String, String) = cmbRcepStyle.SelectedItem
            Dim content = "type:" & kvp.Value & vbCrLf &
                "title:" & chkCustomizeTitle.Checked & vbCrLf &
                "text:" & txtCustomizeTitle.Text
            File.WriteAllText(filePath, content)

            '存過磅單偏移
            WriteRecpMargin(kvp.Key, txtRcepLeft.Text, txtRcepRTop.Text)

            MsgBox("存檔成功")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmbRcepStyle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRcepStyle.SelectedIndexChanged
        Dim cmb As ComboBox = sender
        Dim margins = RoadRecpMargin(cmb.SelectedItem.key)
        txtRcepLeft.Text = margins(0)
        txtRcepRTop.Text = margins(1)
    End Sub

    ''' <summary>
    ''' 更新最後操作的時間
    ''' </summary>
    Private Sub UpdateModifyTime()
        Dim time = Date.Now
        Dim dic = New Dictionary(Of String, String) From {{"更新時間", time.ToString}}
        UpdateTable("資料更新", dic, "編號 = 1")
        tempModify = time
    End Sub

    Private Sub tmrCheckModify_Tick(sender As Object, e As EventArgs) Handles tmrCheckModify.Tick
        Dim time = GetModifyTime()

        If tempModify.ToString <> time.ToString AndAlso Not lblWarningModify.Visible Then
            lblWarningModify.Visible = True
            tempModify = time
        End If
    End Sub

    Private Function GetModifyTime() As Date
        Dim dic = New Dictionary(Of String, Object) From {{"編號", 1}}
        SetCheckTime()
        Return SelectTable("SELECT * FROM 資料更新", dic).Rows(0).Field(Of Date)("更新時間")
    End Function

    Private Sub btnSave_dbCheck_Click(sender As Object, e As EventArgs) Handles btnSave_dbCheck.Click
        Dim dic = New Dictionary(Of String, String) From {{"檢查時間", txtDBCheck.Text}}
        If UpdateTable("資料更新", dic, "編號 = 1") Then
            SetCheckTime()
            MsgBox("儲存成功")
        End If
    End Sub
End Class