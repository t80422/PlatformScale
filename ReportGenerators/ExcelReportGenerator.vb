Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Office.Interop.Excel

Namespace ReportGenerators
    Public Class ExcelReportGenerator
        Private exl As Application
        Private wb As Workbook
        Private ws As Worksheet
        Private cells As Range
        Private _filePath As String
        Private _saveExcel As Boolean

        Public Sub New(filePath As String, Optional saveExcel As Boolean = False)
            _filePath = filePath
            exl = New Application
            _saveExcel = saveExcel
        End Sub

        ''' <summary>
        ''' 年度對帳單
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateYearlyStatement(year As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '寫入標題
            cells(1, 1) = $"{year} 年度 {inOut} 對帳單"

            Dim dic As New Dictionary(Of String, Object) From {
                {"startYear", New Date(year, 1, 1).ToString("yyyy/MM/dd")},
                {"endYear", New Date(year, 12, 31).ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }

            Dim sql = "SELECT 過磅日期, [客戶/廠商], 產品名稱, 車牌號碼, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startYear AND @endYear " &
                      "AND [進/出] = @inOut "

            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            Dim groupData = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    Date.Parse(row("過磅日期")).Month,
                    .Customer = row("客戶/廠商"),
                    .Product = row("產品名稱"),
                    .CarNo = row("車牌號碼")
                }).
                Select(Function(group) New With {
                    group.Key.Month,
                    group.Key.Customer,
                    group.Key.Product,
                    group.Key.CarNo,
                    .Data = group.Select(Function(row) New With {
                        .Weight = Math.Round(row("淨重"), 3),
                        .Meters = Math.Round(row("米數"), 3)
                        }).ToList()
                }).ToList()

            ' 建立 DataTable
            Dim table As New System.Data.DataTable()
            For i As Integer = 1 To 6
                table.Columns.Add()
            Next


            For Each monthGroup In groupData.GroupBy(Function(x) x.Month).OrderBy(Function(x) x.Key)
                '寫入月份
                table.Rows.Add(monthGroup.Key & "月")

                For Each customerGroup In monthGroup.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                    For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                            '寫入          客戶               產品              車號
                            table.Rows.Add(customerGroup.Key, productGroup.Key, CarNoGroup.Key)

                            For Each carNoItem In CarNoGroup
                                For Each record In carNoItem.Data
                                    ' 寫入每條記錄的詳細資料
                                    table.Rows.Add("", "", "", record.Weight, record.Meters)
                                Next
                            Next

                            ' 寫入小計
                            table.Rows.Add("", "", "(小計)", CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Weight))),
                                           CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meters))), CarNoGroup.Count)

                            '' 記住剛剛添加的「小計」行的行號
                            'Dim subtotalRowIndex As Integer = table.Rows.Count

                            '' 添加上邊框到「小計」所在的行
                            'Dim topBorderRange As Range = ws.Range(ws.Cells(subtotalRowIndex + 2, 1), ws.Cells(subtotalRowIndex + 2, table.Columns.Count))
                            'topBorderRange.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            'topBorderRange.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin

                            table.Rows.Add()
                        Next
                    Next
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 4, table.Columns.Count))
            range.Value = objectArray

        End Sub

        ''' <summary>
        ''' 月度對帳單
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyStatement(year As Integer, month As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '撈資料
            Dim startDate As New Date(year, month, 1)
            Dim endDate = startDate.AddMonths(1).AddDays(-1)

            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate.ToString("yyyy/MM/dd")},
                {"endDate", endDate.ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }

            Dim sql = "SELECT [客戶/廠商], 產品代號, 磅單序號, 空重, 總重, 每米噸數, 淨重, 米數, 產品名稱, 車牌號碼 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @endDate AND @startDate " &
                      "AND [進/出] = @inOut "

            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            '標題
            cells(1, 1) = $"{startDate:yyyy年MM月} {inOut} 對帳單"

            '撈取當月客戶購買的產品所使用的車的資料
            Dim datas = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    .Customer = row("客戶/廠商"),
                    .Product = row("產品名稱"),
                    .CarNo = row("車牌號碼")
                }).
                Select(Function(group) New With {
                    group.Key.Customer,
                    group.Key.Product,
                    group.Key.CarNo,
                    .Data = group.Select(Function(row) New With {
                        .ID = row("磅單序號"),
                        .ProductID = row("產品代號"),
                        .Empty = Math.Round(row("空重"), 3),
                        .Weight = Math.Round(row("總重"), 3),
                        .NetWeight = Math.Round(row("淨重"), 3),
                        .Meter = Math.Round(row("米數"), 3),
                        .TPM = row("每米噸數")
                    }).ToList
                }).ToList

            ' 建立 DataTable
            Dim table As New System.Data.DataTable()
            For i As Integer = 1 To 7
                table.Columns.Add()
            Next

            For Each customerGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                '寫入客戶
                table.Rows.Add("客戶:" & customerGroup.Key)

                For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                    '寫入產品
                    table.Rows.Add("產品:" & productGroup.Key)

                    For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                        '寫入車號
                        table.Rows.Add("車號:" & CarNoGroup.Key)
                        table.Rows.Add()
                        For Each carNoItem In CarNoGroup
                            For Each record In carNoItem.Data
                                '寫入每條記錄的詳細資料
                                table.Rows.Add(record.ID, record.ProductID, record.Empty, record.Weight, record.NetWeight, record.Meter, record.TPM)
                            Next
                        Next

                        ' 寫入小計
                        Dim totalNetWeight = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                        Dim totalMeter = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))
                        table.Rows.Add("", "", "", "(小計)", totalNetWeight, totalMeter, CarNoGroup.Count & "輛")
                        table.Rows.Add()
                    Next
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 日報統計明細表
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <param name="startDay"></param>
        ''' <param name="endDay"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyStatement(year As Integer, month As Integer, startDay As Integer, endDay As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '撈資料
            Dim startDate = New Date(year, month, startDay).ToString("yyyy/MM/dd")
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate},
                {"endDate", New Date(year, month, endDay).ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }
            Dim sql = "SELECT 過磅日期, [客戶/廠商], 產品名稱, 車牌號碼, 磅單序號, 空重, 總重, 淨重, 米數, 每米噸數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startDate AND @endDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            '標題
            cells(1, 1) = $"{startDate}~{endDay} {inOut} 產品統計表"

            '撈取當月客戶購買的產品所使用的車的資料
            Dim datas = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    .Day = row("過磅日期"),
                    .Customer = row("客戶/廠商"),
                    .Product = row("產品名稱"),
                    .CarNo = row("車牌號碼")
                }).
                Select(Function(group) New With {
                    group.Key.Day,
                    group.Key.Customer,
                    group.Key.Product,
                    group.Key.CarNo,
                    .Data = group.Select(Function(row) New With {
                        .ID = row("磅單序號"),
                        .Empty = Math.Round(row("空重"), 3),
                        .Weight = Math.Round(row("總重"), 3),
                        .NetWeight = Math.Round(row("淨重"), 3),
                        .Meter = Math.Round(row("米數"), 3),
                        .TPM = row("每米噸數")
                    })
                })

            ' 建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 6
                table.Columns.Add()
            Next

            For Each dayGroup In datas.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                '寫入日期
                table.Rows.Add("日期:" & dayGroup.Key)

                For Each customerGroup In dayGroup.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                    '寫入客戶
                    table.Rows.Add("客戶:" & customerGroup.Key)

                    For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        '寫入產品
                        table.Rows.Add("產品:" & productGroup.Key)

                        For Each CarNoGroup In productGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                            '寫入車號
                            table.Rows.Add("車號:" & CarNoGroup.Key)

                            For Each carNoItem In CarNoGroup
                                For Each record In carNoItem.Data
                                    '寫入每條記錄的詳細資料
                                    table.Rows.Add(record.ID, record.Empty, record.Weight, record.NetWeight, record.Meter, record.TPM)
                                Next
                            Next

                            ' 寫入小計
                            Dim sumNetWeight = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                            Dim sumMeter = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))

                            table.Rows.Add("", "", "(小計)", sumNetWeight, sumMeter, CarNoGroup.Count & "輛")
                        Next
                    Next
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 月報統計表(總量)
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <param name="dayStart"></param>
        ''' <param name="dayEnd"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyReport(year As Integer, month As Integer, dayStart As Integer, dayEnd As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '列標題
            cells(1, 1) = $"{year}年 {month}月 {dayStart}~{dayEnd} {inOut} 月份統計表(總量)"

            '撈資料
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", New Date(year, month, dayStart).ToString("yyyy/MM/dd")},
                {"endDate", New Date(year, month, dayEnd).ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }
            Dim sql = "SELECT [客戶/廠商], 車牌號碼, 產品名稱, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @endDate AND @startDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            '建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 6
                table.Columns.Add()
            Next

            '撈取當月客戶購買的產品所使用的車的資料
            Dim datas = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    .Customer = row("客戶/廠商"),
                    .CarNo = row("車牌號碼"),
                    .Product = row("產品名稱")
                }).
                Select(Function(group) New With {
                    group.Key.Customer,
                    group.Key.CarNo,
                    group.Key.Product,
                    .SumNetWeight = group.Sum(Function(row) Convert.ToDouble(row("淨重"))),
                    .SumMeter = group.Sum(Function(row) Convert.ToDouble(row("米數")))
                }).ToList()

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                table.Rows.Add($"客戶/廠商:{cusGroup.Key}")

                For Each carGroup In cusGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                    table.Rows.Add($"車牌:{carGroup.Key}")

                    Dim totalNetWeight As Double = 0.0
                    Dim totalMeter As Double = 0.0
                    Dim totalCount As Integer = 0

                    For Each prodGroup In carGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        Dim sumNetWeight = Math.Round(prodGroup.First().SumNetWeight, 3)
                        Dim sumMeter = Math.Round(prodGroup.First().SumMeter, 3)
                        table.Rows.Add(prodGroup.Key, prodGroup.Count, sumNetWeight, sumMeter)

                        totalNetWeight += sumNetWeight
                        totalMeter += sumMeter
                        totalCount += prodGroup.Count
                    Next

                    table.Rows.Add("(小計)", totalCount, totalNetWeight, totalMeter)
                    table.Rows.Add()
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 月報統計表(產品)
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyProductStats(year As Integer, month As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            cells(1, 1) = $"{year}年{month}月 產品 {inOut} 統計表"

            '撈資料
            Dim sql = "SELECT 產品名稱, COUNT(*) AS 車次, SUM(總重) AS 總重, SUM(米數) AS 米 FROM 過磅資料表 " &
                      $"WHERE YEAR(過磅日期) = {year} " &
                      $"AND MONTH(過磅日期) = {month} " &
                      $"AND [進/出] = '{inOut}' "

            sql = AppendConditionsToSQL(sql, dicCondition) &
                  "GROUP BY 產品名稱 " &
                  "ORDER BY 產品名稱"

            Dim dt = SelectTable(sql)

            Dim rowIndex = 3

            Dim sumCarCount As Integer = 0
            Dim sumWeight As Double = 0
            Dim sumMeter As Double = 0
            'todo 改善寫入速度
            For Each row As DataRow In dt.Rows
                cells(rowIndex, 1) = row("產品名稱")
                cells(rowIndex, 2) = row("車次")
                sumCarCount += row("車次")
                cells(rowIndex, 3) = Math.Round(row("總重"), 3)
                sumWeight += row("總重")
                cells(rowIndex, 4) = Math.Round(row("米"), 3)
                sumMeter += row("米")

                rowIndex += 1
            Next

            For i As Integer = 1 To 5
                TopLine_Cell(cells(rowIndex, i))
            Next

            cells(rowIndex, 1) = "(總計)"
            cells(rowIndex, 2) = Math.Round(sumCarCount, 3)
            cells(rowIndex, 3) = Math.Round(sumWeight, 3)
            cells(rowIndex, 4) = Math.Round(sumMeter, 3)
        End Sub

        ''' <summary>
        ''' 日報統計表(產品)
        ''' </summary>
        ''' <param name="startDate"></param>
        ''' <param name="endDate"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyProductStats(startDate As String, endDate As String, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            cells(1, 1) = $"產品每月(日) {inOut} 統計表"

            '撈資料
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate},
                {"endDate", endDate},
                {"inOut", inOut}
            }
            Dim sql = "SELECT 過磅日期, 產品名稱, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startDate AND @endDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            '分類資料
            Dim datas = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    .Day = row("過磅日期"),
                    .Product = row("產品名稱")
                }).
                Select(Function(group) New With {
                    group.Key.Day,
                    group.Key.Product,
                    .SumNetWeight = Math.Round(group.Sum(Function(row) Double.Parse(row("淨重"))), 3),
                    .SumMeter = Math.Round(group.Sum(Function(row) Double.Parse(row("米數"))), 3)
                })

            '建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 6
                table.Columns.Add()
            Next

            For Each dayGroup In datas.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                table.Rows.Add("日期:" & dayGroup.Key)

                Dim totalNetWeight As Double = 0.0
                Dim totalMeter As Double = 0.0
                Dim totalCount As Integer = 0

                For Each prodGroup In dayGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                    Dim sumNetWeight = prodGroup.First().SumNetWeight
                    Dim sumMeter = prodGroup.First().SumMeter
                    table.Rows.Add(prodGroup.Key, prodGroup.Count, sumNetWeight, sumMeter)

                    totalNetWeight += sumNetWeight
                    totalMeter += sumMeter
                    totalCount += prodGroup.Count
                Next

                table.Rows.Add("(小計)", totalCount, totalNetWeight, totalMeter)
                table.Rows.Add()
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 日統計明細表(產品 客戶)
        ''' </summary>
        ''' <param name="startDate"></param>
        ''' <param name="endDate"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyProductCustomerStats(startDate As String, endDate As String, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            Dim person As String = ""

            Select Case inOut
                Case "進貨"
                    person = "廠商"
                Case "出貨"
                    person = "客戶"
                Case Else

            End Select

            cells(1, 1) = $"{person}每日 {inOut} 產品統計表"

            '撈資料
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate},
                {"endDate", endDate},
                {"inOut", inOut}
            }
            Dim sql = "SELECT [客戶/廠商], 過磅日期, 產品名稱, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startDate AND @endDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            Dim datas = dt.AsEnumerable.
                GroupBy(Function(row) New With {
                    .Customer = row("客戶/廠商"),
                    .Day = row("過磅日期")
                }).
                Select(Function(group) New With {
                    group.Key.Customer,
                    group.Key.Day,
                    .Data = group.Select(Function(row) New With {
                        .Product = row("產品名稱"),
                        .NetWeight = Math.Round(row("淨重"), 3),
                        .Meter = Math.Round(row("米數"), 3)
                    })
                })

            ' 建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 3
                table.Columns.Add()
            Next

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                table.Rows.Add("客戶:" & cusGroup.Key)

                For Each dayGroup In cusGroup.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                    table.Rows.Add("日期:" & dayGroup.Key)

                    For Each dayItem In dayGroup
                        For Each d In dayItem.Data
                            table.Rows.Add(d.Product, d.NetWeight, d.Meter)
                        Next
                    Next

                    ' 寫入小計
                    Dim sumNetWeight = dayGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                    Dim sumMeter = dayGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))

                    table.Rows.Add("(小計)", sumNetWeight, sumMeter)
                    table.Rows.Add()
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 日報統計表(客戶 產品)
        ''' </summary>
        ''' <param name="startDate"></param>
        ''' <param name="endDate"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyCustomerProductStats(startDate As String, endDate As String, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            Dim person As String = ""

            Select Case inOut
                Case "進貨"
                    person = "廠商"
                Case "出貨"
                    person = "客戶"
                Case Else

            End Select

            cells(1, 1) = $"{person}每日 {inOut} 產品統計表"

            '撈資料
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate},
                {"endDate", endDate},
                {"inOut", inOut}
            }
            Dim sql = "SELECT [客戶/廠商], 過磅日期, 產品名稱, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startDate AND @endDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            '分類資料
            Dim datas = dt.AsEnumerable().
                GroupBy(Function(row) New With {
                    .Customer = row("客戶/廠商"),
                    .Day = row("過磅日期"),
                    .Product = row("產品名稱")
                }).
                Select(Function(group) New With {
                    group.Key.Customer,
                    group.Key.Day,
                    group.Key.Product,
                    .SumNetWeight = Math.Round(group.Sum(Function(row) Double.Parse(row("淨重"))), 3),
                    .SumMeter = Math.Round(group.Sum(Function(row) Double.Parse(row("米數"))), 3)
                })

            '建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 6
                table.Columns.Add()
            Next

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                '列出客戶
                table.Rows.Add("客戶:" & cusGroup.Key)

                For Each dayGroup In cusGroup.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                    '列出日期
                    table.Rows.Add("日期:" & dayGroup.Key)

                    Dim totalCount As Integer = 0
                    Dim totalNetWeight As Double = 0
                    Dim totalMeter As Double = 0

                    For Each prodGroup In dayGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        Dim sumNetWeight = prodGroup.First().SumNetWeight
                        Dim sumMeter = prodGroup.First().SumMeter
                        table.Rows.Add(prodGroup.Key, prodGroup.Count, sumNetWeight, sumMeter)

                        totalNetWeight += sumNetWeight
                        totalMeter += sumMeter
                        totalCount += prodGroup.Count
                    Next

                    table.Rows.Add("(小計)", totalCount, totalNetWeight, totalMeter)
                    table.Rows.Add()
                Next
            Next

            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 3, table.Columns.Count))
            range.Value = objectArray
        End Sub

        ''' <summary>
        ''' 過磅單日統計表
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <param name="strartDay"></param>
        ''' <param name="endDay"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateWeighingDailyReport(year As Integer, month As Integer, startDay As Integer, endDay As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            cells(1, 1) = $"{year}年{month}月{startDay}~{endDay} {inOut} 每日過磅明細表"

            '撈資料
            Dim startDate = New Date(year, month, startDay).ToString("yyyy/MM/dd")
            Dim endDate = New Date(year, month, endDay).ToString("yyyy/MM/dd")
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate},
                {"endDate", endDate},
                {"inOut", inOut}
            }
            Dim sql = "SELECT 磅單序號, [客戶/廠商], 車牌號碼, 產品名稱, 總重, 空重, 淨重, 米數 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @startDate AND @endDate " &
                      "AND [進/出] = @inOut "
            Dim dt = SelectTable(AppendConditionsToSQL(sql, dicCondition), dic)

            Dim rowIndex = 3

            Dim sumWeight As Double = 0
            Dim sumEmptyWeight As Double = 0
            Dim sumNetWeight As Double = 0
            Dim sumMeter As Double = 0

            'todo 改善寫入速度
            For Each row As DataRow In dt.Rows
                Dim weight As Double = Math.Round(row("總重"), 3)
                Dim emptyWeight As Double = Math.Round(row("空重"), 3)
                Dim netWeight As Double = Math.Round(row("淨重"), 3)
                Dim meter As Double = Math.Round(row("米數"), 3)

                cells(rowIndex, 1) = row("磅單序號")
                cells(rowIndex, 2) = row("客戶/廠商")
                cells(rowIndex, 3) = row("車牌號碼")
                cells(rowIndex, 4) = row("產品名稱")
                cells(rowIndex, 5) = weight
                sumWeight += weight
                cells(rowIndex, 6) = emptyWeight
                sumEmptyWeight += emptyWeight
                cells(rowIndex, 7) = netWeight
                sumNetWeight += netWeight
                cells(rowIndex, 8) = meter
                sumMeter += meter

                rowIndex += 1
            Next

            cells(rowIndex, 4) = "(總計)"
            cells(rowIndex, 5) = sumWeight
            cells(rowIndex, 6) = sumEmptyWeight
            cells(rowIndex, 7) = sumNetWeight
            cells(rowIndex, 8) = sumMeter
        End Sub

        Public Sub CreateNewReport(sheetName As String)
            Dim orgWb As Workbook = exl.Workbooks.Open(Path.Combine(_filePath, "Report", "報表範本檔.xlsx"))
            Dim orgWs As Worksheet = orgWb.Worksheets(sheetName)
            wb = exl.Workbooks.Add
            orgWs.Copy(wb.Sheets(1))
            orgWb.Close(False)
            wb.Sheets(2).Delete
            ws = wb.Worksheets(sheetName)
            cells = ws.Cells
        End Sub

        Public Sub SaveReport(sheet As String)
            Dim pdfFilePath = Path.Combine(_filePath, "Report", $"{sheet}.pdf")
            Dim exlFilePath = Path.Combine(_filePath, "Report", $"{sheet}.xlsx")

            If _saveExcel Then ws.SaveAs(exlFilePath)

            ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFilePath)
            Process.Start(pdfFilePath)
            Close()
        End Sub

        Private Sub ReleaseObject(obj As Object)
            Try
                Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End Sub

        ''' <summary>
        ''' 關閉Excel應用程序和任何打開的工作簿。
        ''' </summary>
        ''' <param name="saveChanges">指定是否在關閉前保存工作簿的更改。</param>
        Public Sub Close(Optional saveChanges As Boolean = False)
            Try
                ' 關閉工作簿
                If wb IsNot Nothing Then
                    wb.Close(saveChanges)
                    ReleaseObject(wb)
                    wb = Nothing
                End If

                ' 關閉Excel應用程序
                If exl IsNot Nothing Then
                    exl.Quit()
                    ReleaseObject(exl)
                    exl = Nothing
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 將儲存格加上細的下框線
        ''' </summary>
        ''' <param name="row"></param>
        ''' <param name="colStart"></param>
        ''' <param name="colEnd"></param>
        Protected Sub BottomLine_Cell(row As Integer, colStart As Integer, colEnd As Integer)
            For i = colStart To colEnd
                cells(row, i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                cells(row, i).Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
            Next
        End Sub

        ''' <summary>
        ''' 將儲存格加上細的上框線
        ''' </summary>
        ''' <param name="cell"></param>
        Protected Sub TopLine_Cell(cell As Range)
            cell.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            cell.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
        End Sub

        Protected Sub TopLine_Cell(row As Integer, colStart As Integer, colEnd As Integer)
            For i = colStart To colEnd
                cells(row, i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                cells(row, i).Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin

            Next
        End Sub

        Private Function AppendConditionsToSQL(sql As String, dicCondition As Dictionary(Of String, String)) As String
            Dim sb As New StringBuilder(sql)

            For Each kvp In dicCondition
                If kvp.Value <> "全部" Then sb.Append($"AND {kvp.Key} = '{kvp.Value}' ")
            Next

            Return sb.ToString
        End Function
    End Class
End Namespace