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

            Dim datas = From row In dt
                        Group By
                            Date.Parse(row("過磅日期")).Month,
                            Customer = row("客戶/廠商"),
                            Product = row("產品名稱"),
                            CarNo = row("車牌號碼")
                        Into Group
                        Select New With {
                            Key Month,
                            Key Customer,
                            Key Product,
                            Key CarNo,
                            Group.Count,
                            .SumNetWeight = Group.Sum(Function(row) Double.Parse(row("淨重"))),
                            .SumMeter = Group.Sum(Function(row) Double.Parse(row("米數")))
                         }

            ' 建立 DataTable
            Using table As New Data.DataTable()

                For i As Integer = 1 To 4
                    table.Columns.Add()
                Next

                Dim yearNetWeight As Double
                Dim yearMeter As Double
                Dim yearCount As Integer
                Dim rowIndex As Integer = 3

                For Each monthGroup In datas.GroupBy(Function(x) x.Month).OrderBy(Function(x) x.Key)
                    '寫入月份
                    table.Rows.Add(monthGroup.Key & "月")
                    DrawLine(rowIndex, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)
                    rowIndex += 1

                    Dim monthNetWeight As Double = 0.0
                    Dim monthMeter As Double = 0.0
                    Dim monthCount As Integer = 0

                    For Each customerGroup In monthGroup.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                        table.Rows.Add("客戶/廠商:" & customerGroup.Key)
                        DrawLine(rowIndex, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)
                        rowIndex += 1

                        Dim totalNetWeight As Double = 0.0
                        Dim totalMeter As Double = 0.0
                        Dim totalCount As Integer = 0

                        For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                            table.Rows.Add("產品:" & productGroup.Key)
                            DrawLine(rowIndex, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)
                            rowIndex += 1

                            Dim prodNetWeight As Double = 0.0
                            Dim prodMeter As Double = 0.0
                            Dim prodCount As Integer = 0

                            For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                                table.Rows.Add("車號:" & CarNoGroup.Key)
                                DrawLine(rowIndex, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)
                                rowIndex += 1

                                table.Rows.Add("(車號小計)", CarNoGroup.First.SumNetWeight, CarNoGroup.First.SumMeter, CarNoGroup.First.Count)
                                DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                                rowIndex += 1

                                prodNetWeight += CarNoGroup.First.SumNetWeight
                                prodMeter += CarNoGroup.First.SumMeter
                                prodCount += CarNoGroup.First.Count
                            Next

                            table.Rows.Add("(產品小計)", prodNetWeight, prodMeter, prodCount)
                            DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                            rowIndex += 1
                            totalNetWeight += prodNetWeight
                            totalMeter += prodMeter
                            totalCount += prodCount
                        Next

                        table.Rows.Add("(客戶小計)", totalNetWeight, totalMeter, totalCount)
                        DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                        rowIndex += 1
                        monthNetWeight += totalNetWeight
                        monthMeter += totalMeter
                        monthCount += totalCount
                    Next

                    table.Rows.Add("(月份小計)", monthNetWeight, monthMeter, monthCount)
                    DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                    rowIndex += 1
                    yearNetWeight += monthNetWeight
                    yearMeter += monthMeter
                    yearCount += monthCount
                Next

                table.Rows.Add("(總計)", yearNetWeight, yearMeter, yearCount)

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
            End Using
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

            Dim datas = From row In dt
                        Group By
                            Customer = row("客戶/廠商"),
                            Product = row("產品名稱"),
                            CarNo = row("車牌號碼")
                        Into Group
                        Select New With {
                            Key Customer,
                            Key Product,
                            Key CarNo,
                            Group.Count,
                            .SumNetWeight = Group.Sum(Function(row) Double.Parse(row("淨重"))),
                            .SumMeter = Group.Sum(Function(row) Double.Parse(row("米數"))),
                            .Data = From row In Group
                                    Select New With {
                                .ID = row("磅單序號"),
                                .ProductID = row("產品代號"),
                                .Empty = Math.Round(row("空重"), 3),
                                .Weight = Math.Round(row("總重"), 3),
                                .NetWeight = Math.Round(row("淨重"), 3),
                                .Meter = Math.Round(row("米數"), 3),
                                .TPM = row("每米噸數")
                            }
                        }

            ' 建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 7
                table.Columns.Add()
            Next

            Dim monthNetWeight As Double = 0.0
            Dim monthMeter As Double = 0.0
            Dim monthCount As Integer = 0

            For Each customerGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                Dim cusNet As Double = 0
                Dim cusMeter As Double = 0
                Dim cusCount As Integer = 0

                table.Rows.Add("客戶:" & customerGroup.Key)
                DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                    Dim prodNet As Double = 0
                    Dim prodMeter As Double = 0
                    Dim prodCount As Integer = 0

                    table.Rows.Add("產品:" & productGroup.Key)
                    DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                    For Each CarNoGroup In productGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                        table.Rows.Add("車號:" & CarNoGroup.Key)
                        DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                        For Each carNoItem In CarNoGroup
                            For Each record In carNoItem.Data
                                '寫入每條記錄的詳細資料
                                table.Rows.Add(record.ID, record.ProductID, record.Empty, record.Weight, record.NetWeight, record.Meter, record.TPM)
                                DrawLine(table.Rows.Count + 2, 1, 7, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                            Next
                        Next

                        table.Rows.Add("", "", "", "(車號小計)", CarNoGroup.First.SumNetWeight, CarNoGroup.First.SumMeter, CarNoGroup.First.Count & "輛")
                        DrawLine(table.Rows.Count + 2, 4, 7, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)

                        prodNet += CarNoGroup.First.SumNetWeight
                        prodMeter += CarNoGroup.First.SumMeter
                        prodCount += CarNoGroup.First.Count
                    Next

                    table.Rows.Add("", "", "", "(產品小計)", prodNet, prodMeter, prodCount & "輛")
                    DrawLine(table.Rows.Count + 2, 4, 7, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)
                    cusNet += prodNet
                    cusMeter += prodMeter
                    cusCount += prodCount
                Next

                table.Rows.Add("", "", "", "(客戶小計)", cusNet, cusMeter, cusCount & "輛")
                DrawLine(table.Rows.Count + 2, 4, 7, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)
                monthNetWeight += cusNet
                monthMeter += cusMeter
                monthCount += cusCount
            Next

            table.Rows.Add("", "", "", "(總計)", monthNetWeight, monthMeter, monthCount & "輛")
            DrawLine(table.Rows.Count + 2, 4, 7, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)
            WriteToExcel(table)
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

            Dim totalNetWeight As Double = 0
            Dim totalMeter As Double = 0
            Dim totalCount As Integer = 0

            For Each dayGroup In datas.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                table.Rows.Add("日期:" & dayGroup.Key)
                DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                For Each customerGroup In dayGroup.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                    table.Rows.Add("客戶:" & customerGroup.Key)
                    DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                    For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        table.Rows.Add("產品:" & productGroup.Key)
                        DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                        For Each CarNoGroup In productGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                            table.Rows.Add("車號:" & CarNoGroup.Key)
                            DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                            For Each carNoItem In CarNoGroup
                                For Each record In carNoItem.Data
                                    '寫入每條記錄的詳細資料
                                    table.Rows.Add(record.ID, record.Empty, record.Weight, record.NetWeight, record.Meter, record.TPM)
                                    DrawLine(table.Rows.Count + 2, 1, 6, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                                Next
                            Next

                            ' 寫入小計
                            Dim sumNetWeight = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                            Dim sumMeter = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))

                            table.Rows.Add("", "", "(小計)", sumNetWeight, sumMeter, CarNoGroup.Count & "輛")
                            totalNetWeight += sumNetWeight
                            totalMeter += sumMeter
                            totalCount += CarNoGroup.Count
                        Next
                    Next
                Next
            Next

            DrawLine(table.Rows.Count + 2, 1, 6, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
            table.Rows.Add("", "", "(總計)", totalNetWeight, totalMeter, totalCount & "輛")

            WriteToExcel(table)
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

            For i As Integer = 1 To 4
                table.Columns.Add()
            Next

            Dim datas = From row In dt
                        Group By
                            Customer = row("客戶/廠商"),
                            CarNo = row("車牌號碼"),
                            Product = row("產品名稱")
                        Into Group
                        Select New With {
                            Key Customer,
                            Key CarNo,
                            Key Product,
                            Group.Count,
                            .SumMeter = Group.Sum(Function(row) Double.Parse(row("米數"))),
                            .SumNetWeight = Group.Sum(Function(row) Double.Parse(row("淨重")))
                        }

            Dim totalNetWeight As Double = 0.0
            Dim totalMeter As Double = 0.0
            Dim totalCount As Integer = 0

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                table.Rows.Add($"客戶/廠商:{cusGroup.Key}")
                DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                For Each carGroup In cusGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                    table.Rows.Add($"車牌:{carGroup.Key}")
                    DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                    Dim sumNetWeight As Double = 0.0
                    Dim sumMeter As Double = 0.0
                    Dim sumCount As Integer = 0

                    For Each prodGroup In carGroup
                        table.Rows.Add(prodGroup.Product, prodGroup.Count, prodGroup.SumNetWeight, prodGroup.SumMeter)
                        DrawLine(table.Rows.Count + 2, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)

                        sumNetWeight += prodGroup.SumNetWeight
                        sumMeter += prodGroup.SumMeter
                        sumCount += prodGroup.Count
                    Next

                    table.Rows.Add("(小計)", sumCount, sumNetWeight, sumMeter)
                    table.Rows.Add()

                    totalNetWeight += sumNetWeight
                    totalMeter += sumMeter
                    totalCount += sumCount
                Next
            Next

            table.Rows.RemoveAt(table.Rows.Count - 1)
            DrawLine(table.Rows.Count + 2, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
            table.Rows.Add("(總計)", totalCount, totalNetWeight, totalMeter)

            WriteToExcel(table)
        End Sub

        ''' <summary>
        ''' 月報統計表(產品)
        ''' </summary>
        ''' <param name="year"></param>
        ''' <param name="month"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyProductStats(year As Integer, month As Integer, dayStart As Integer, dayEnd As Integer, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            cells(1, 1) = $"{year}年{month}月 產品 {inOut} 統計表"

            cells(2, 1) = $"客戶:{dicCondition("[客戶/廠商]")}"
            cells(3, 1) = $"車號:{dicCondition("車牌號碼")}"

            '撈資料
            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", New Date(year, month, dayStart).ToString("yyyy/MM/dd")},
                {"endDate", New Date(year, month, dayEnd).ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }
            Dim sql = "SELECT 產品名稱, COUNT(*) AS 車次, SUM(淨重) AS 總重, SUM(米數) AS 米 FROM 過磅資料表 " &
                      "WHERE 過磅日期 BETWEEN @endDate AND @startDate " &
                      $"AND [進/出] = @inOut "

            sql = AppendConditionsToSQL(sql, dicCondition) &
                  "GROUP BY 產品名稱 " &
                  "ORDER BY 產品名稱"

            Dim dt = SelectTable(sql, dic)

            Dim rowIndex = 5

            Dim dataArray(dt.Rows.Count, 4) As Object

            Dim sumCarCount As Integer = 0
            Dim sumWeight As Double = 0
            Dim sumMeter As Double = 0

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim row As DataRow = dt.Rows(i)
                dataArray(i, 0) = row("產品名稱")
                dataArray(i, 1) = row("車次")
                sumCarCount += row("車次")
                dataArray(i, 2) = Math.Round(row("總重"), 3)
                sumWeight += row("總重")
                dataArray(i, 3) = Math.Round(row("米"), 3)
                sumMeter += row("米")
                DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                rowIndex += 1
            Next

            cells.Range("A5").Resize(dt.Rows.Count, 4).Value = dataArray

            DrawLine(rowIndex, 1, 4, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)
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

            Dim datas = From row In dt
                        Group By
                            Day = row("過磅日期"),
                            Product = row("產品名稱")
                        Into Group
                        Select New With {
                            Key Day,
                            Key Product,
                            Group.Count,
                            .SumNetWeight = Group.Sum(Function(x) Double.Parse(x("淨重"))),
                            .SumMeter = Group.Sum(Function(x) Double.Parse(x("米數")))
                        }

            '建立 DataTable
            Dim table As New Data.DataTable()

            For i As Integer = 1 To 5
                table.Columns.Add()
            Next

            Dim sumNet As Double
            Dim sumMeter As Double
            Dim sumCount As Integer

            For Each dayGroup In datas.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                table.Rows.Add($"日期:{dayGroup.Key}")
                DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                Dim totalNetWeight As Double = 0.0
                Dim totalMeter As Double = 0.0
                Dim totalCount As Integer = 0

                For Each prodGroup In dayGroup
                    table.Rows.Add($"", prodGroup.Product, prodGroup.Count, prodGroup.SumNetWeight, prodGroup.SumMeter)
                    DrawLine(table.Rows.Count + 2, 2, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                    totalCount += prodGroup.Count
                    totalNetWeight += prodGroup.SumNetWeight
                    totalMeter += prodGroup.SumMeter
                Next

                table.Rows.Add($"", "", totalCount, totalNetWeight, totalMeter)
                table.Rows.Add()

                sumCount += totalCount
                sumNet += totalNetWeight
                sumMeter += totalMeter
            Next

            table.Rows.Add($"", "(總計)", sumCount, sumNet, sumMeter)
            DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeTop, XlBorderWeight.xlThin)

            WriteToExcel(table)
        End Sub

        ''' <summary>
        ''' 日統計明細表(產品 客戶)
        ''' </summary>
        ''' <param name="startDate"></param>
        ''' <param name="endDate"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyProductCustomerStats(startDate As String, endDate As String, inOut As String, dicCondition As Dictionary(Of String, String))
            '標題
            Dim person As String = If(inOut = "進貨", "廠商", If(inOut = "出貨", "客戶", String.Empty))

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

            Dim datas = From row In dt
                        Group By
                            Customer = row("客戶/廠商"),
                            Day = row("過磅日期"),
                            Product = row("產品名稱")
                        Into Group
                        Select New With {
                            Key Customer,
                            Key Day,
                            Key Product,
                            Group.Count,
                            .sumNetWeight = Group.Sum(Function(r) Double.Parse(r("淨重"))),
                            .sumMeter = Group.Sum(Function(r) Double.Parse(r("米數")))
                        }

            ' 建立 DataTable
            Dim table As New Data.DataTable()

            For i As Integer = 1 To 6
                table.Columns.Add()
            Next

            Dim totalCount As Integer = 0
            Dim totalNet As Double = 0
            Dim totalMeter As Double = 0

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                table.Rows.Add("客戶:" & cusGroup.Key)
                DrawLine(table.Rows.Count + 2, 1, 1, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                Dim sumCount As Integer = cusGroup.Sum(Function(x) x.Count)
                Dim sumNetWeight As Double = cusGroup.Sum(Function(x) x.sumNetWeight)
                Dim sumMeter As Double = cusGroup.Sum(Function(x) x.sumMeter)

                For Each dayGroup In cusGroup.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                    '日期
                    table.Rows.Add(dayGroup.Key)

                    For Each prodGroup In dayGroup
                        '                  品名               車次             總淨重                  總米數
                        table.Rows.Add("", prodGroup.Product, prodGroup.Count, prodGroup.sumNetWeight, prodGroup.sumMeter)
                        DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
                    Next
                Next

                ' 寫入小計
                table.Rows.Add("", "(小計)", sumCount, sumNetWeight, sumMeter)
                table.Rows.Add()

                totalCount += sumCount
                totalNet += sumNetWeight
                totalMeter += sumMeter
            Next

            '寫入總計
            DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThick)
            table.Rows.Add("", "(總計)", totalCount, totalNet, totalMeter)

            WriteToExcel(table)
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
            Dim datas = From row In dt
                        Group By
                            Customer = row("客戶/廠商"),
                            Day = row("過磅日期"),
                            Product = row("產品名稱")
                        Into Group
                        Select New With {
                            Key Customer,
                            Key Day,
                            Key Product,
                            Group.Count,
                            .SumNetWeight = Group.Sum(Function(row) Double.Parse(row("淨重"))),
                            .SumMeter = Group.Sum(Function(row) Double.Parse(row("米數")))
            }

            '建立 DataTable
            Dim table As New Data.DataTable()

            For i As Integer = 1 To 5
                table.Columns.Add()
            Next

            For Each cusGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                Dim totalCount As Integer = 0
                Dim totalNet As Double = 0
                Dim totalMeter As Double = 0

                table.Rows.Add("客戶:" & cusGroup.Key)
                DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)

                For Each dayGroup In cusGroup.GroupBy(Function(x) x.Day).OrderBy(Function(x) x.Key)
                    table.Rows.Add(dayGroup.Key)

                    For Each prodGroup In dayGroup
                        table.Rows.Add("", prodGroup.Product, prodGroup.Count, prodGroup.SumNetWeight, prodGroup.SumMeter)
                        DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)

                        totalNet += prodGroup.SumNetWeight
                        totalMeter += prodGroup.SumMeter
                        totalCount += prodGroup.Count
                    Next
                Next

                DrawLine(table.Rows.Count + 2, 1, 5, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlThin)
                table.Rows.Add("總計", "", totalCount, totalNet, totalMeter)
                table.Rows.Add()
            Next

            WriteToExcel(table)
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
            Dim sumWeight As Double
            Dim sumEmptyWeight As Double
            Dim sumNetWeight As Double
            Dim sumMeter As Double

            '建立 DataTable
            Dim table As New Data.DataTable()
            For i As Integer = 1 To 8
                table.Columns.Add()
            Next

            For Each row As DataRow In dt.Rows
                Dim weight As Double = Math.Round(row("總重"), 3)
                Dim emptyWeight As Double = Math.Round(row("空重"), 3)
                Dim netWeight As Double = Math.Round(row("淨重"), 3)
                Dim meter As Double = row("米數")

                table.Rows.Add(row("磅單序號"), row("客戶/廠商"), row("車牌號碼"), row("產品名稱"), weight, emptyWeight, netWeight, "'" & meter.ToString("F2"))
                sumWeight += weight
                sumEmptyWeight += emptyWeight
                sumNetWeight += netWeight
                sumMeter += meter
            Next

            DrawLine(table.Rows.Count + 2, 1, 8, XlBordersIndex.xlEdgeBottom, XlBorderWeight.xlHairline)
            table.Rows.Add("", "(總計)", dt.Rows.Count, "", sumWeight, sumEmptyWeight, sumNetWeight, "'" & Math.Round(sumMeter, 3).ToString("F2"))

            WriteToExcel(table)
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

        Private Sub DrawLine(row As Integer, startCol As Integer, endCol As Integer, borderSide As XlBordersIndex, borderWeight As XlBorderWeight)
            Dim topBorderRange As Range = ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
            topBorderRange.Borders(borderSide).LineStyle = XlLineStyle.xlContinuous
            topBorderRange.Borders(borderSide).Weight = borderWeight
        End Sub

        Private Function AppendConditionsToSQL(sql As String, dicCondition As Dictionary(Of String, String)) As String
            Dim sb As New StringBuilder(sql)

            For Each kvp In dicCondition
                If kvp.Value <> "全部" Then sb.Append($"AND {kvp.Key} = '{kvp.Value}' ")
            Next

            Return sb.ToString
        End Function


        Public Sub WriteToExcel(table As Data.DataTable)
            ' 將 DataTable 寫入一個二維陣列
            Dim objectArray(table.Rows.Count, table.Columns.Count - 1) As Object

            ' 寫入資料
            For i = 0 To table.Rows.Count - 1
                For j = 0 To table.Columns.Count - 1
                    objectArray(i, j) = table.Rows(i)(j)
                Next
            Next

            ' 一次性將二維陣列寫入 Excel
            Dim range As Range = ws.Range(ws.Cells(3, 1), ws.Cells(table.Rows.Count + 2, table.Columns.Count))
            range.Value = objectArray
        End Sub
    End Class
End Namespace