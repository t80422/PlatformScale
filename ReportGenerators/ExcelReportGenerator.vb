Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Namespace ReportGenerators
    Public Class ExcelReportGenerator
        Private exl As Application
        Private wb As Workbook
        Private ws As Worksheet
        Private cells As Range
        Private templatePath As String = "C:\Users\t8042\Desktop\報表.xlsx"
        Private savePath As String = "C:\Users\t8042\Desktop\test.xlsx"

        Public Sub New()
            exl = New Application
        End Sub

        'Public Sub GenerateYearlyStatement(dtp As DateTimePicker, inOut As String, a As String)
        '    Dim year = dtp.Value.Year
        '    Dim dt = SelectTable(
        '        "SELECT 過磅日期, [客戶/廠商], 產品名稱, 車牌號碼, 淨重, 米數, 總價 FROM 過磅資料表 " &
        '       $"WHERE YEAR(CDATE(過磅日期)) = '{year}' " &
        '       $"AND [進/出] = '{inOut}'"
        '                    )

        '    '寫入Excel
        '    '標題
        '    cells(1, 1) = $"{year} 年度 {inOut} 對帳單"

        '    '撈所有月份
        '    Dim monthes = (
        '        From row In dt
        '        Select Date.Parse(row("過磅日期")).Month
        '    ).Distinct

        '    Dim rowIndex = 3

        '    For Each m In monthes
        '        rowIndex = YearlyData(dt, m, cells, rowIndex)
        '    Next
        'End Sub

        'Private Function YearlyData(dt As Data.DataTable, month As Integer, cells As Range, startRowIndex As Integer) As Integer
        '    ' 月份
        '    cells(startRowIndex, 1) = month & " 月"
        '    BottomLine_Cell(cells(startRowIndex, 1))

        '    Dim rowIndex = startRowIndex + 1

        '    ' 撈取當月客戶購買的產品所使用的車的資料
        '    Dim datas = dt.AsEnumerable().
        '    Where(Function(row) Date.Parse(row("過磅日期")).Month = month).
        '    GroupBy(Function(row) New With {
        '        .Customer = row("客戶/廠商"),
        '        .Product = row("產品名稱"),
        '        .LicensePlate = row("車牌號碼")
        '    }).
        '    Select(Function(group) New With {
        '        group.Key.Customer,
        '        group.Key.Product,
        '        group.Key.LicensePlate,
        '        .Data = group.Select(Function(row) New With {
        '            .Weight = Math.Round(row("淨重"), 3),
        '            .Meters = Math.Round(row("米數"), 3),
        '            .Price = Math.Round(row("總價"), 3)
        '        })
        '    })

        '    For Each item In datas
        '        ' 客戶、產品、車號
        '        cells(rowIndex, 1) = item.Customer
        '        cells(rowIndex, 2) = item.Product
        '        cells(rowIndex, 3) = item.LicensePlate
        '        BottomLine_Cell(cells(rowIndex, 1))
        '        BottomLine_Cell(cells(rowIndex, 2))
        '        BottomLine_Cell(cells(rowIndex, 3))
        '        rowIndex += 1

        '        For Each d In item.Data
        '            ' 資料
        '            Dim wt = d.Weight
        '            Dim mt = d.Meters
        '            Dim pc = d.Price

        '            cells(rowIndex, 4).Value = wt
        '            cells(rowIndex, 5).Value = mt
        '            cells(rowIndex, 6).Value = pc
        '            rowIndex += 1
        '        Next

        '        ' 小計
        '        cells(rowIndex, 3).Value = "(小計)"
        '        cells(rowIndex, 4).Value = item.Data.Sum(Function(d) d.Weight)
        '        TopLine_Cell(cells(rowIndex, 4))
        '        cells(rowIndex, 5).Value = item.Data.Sum(Function(d) d.Meters)
        '        TopLine_Cell(cells(rowIndex, 5))
        '        cells(rowIndex, 6).Value = item.Data.Sum(Function(d) d.Price)
        '        TopLine_Cell(cells(rowIndex, 6))
        '        cells(rowIndex, 7).Value = item.Data.Count()
        '        TopLine_Cell(cells(rowIndex, 7))

        '        rowIndex += 2
        '    Next

        '    Return rowIndex
        'End Function

        ''' <summary>
        ''' 年度對帳單
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateYearlyStatement(dtp As DateTimePicker, inOut As String)
            Dim year = dtp.Value.Year

            '寫入標題
            cells(1, 1) = $"{year} 年度 {inOut} 對帳單"

            Dim dic As New Dictionary(Of String, Object) From {
                {"startYear", New Date(year, 1, 1).ToString("yyyy/MM/dd")},
                {"endYear", New Date(year, 12, 31).ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }

            Dim dt = SelectTable(
                "SELECT 過磅日期, [客戶/廠商], 產品名稱, 車牌號碼, 淨重, 米數, 總價 FROM 過磅資料表 " &
                "WHERE 過磅日期 BETWEEN @startYear AND @endYear " &
                "AND [進/出] = @inOut",
                dic
            )

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
                        .Meters = Math.Round(row("米數"), 3),
                        .Price = Math.Round(row("總價"), 3)
                        }).ToList()
                }).ToList()

            'Dim rowIndex = 3

            ' 建立 DataTable
            Dim table As New Data.DataTable()
            With table
                .Columns.Add("客戶")
                .Columns.Add("產品")
                .Columns.Add("車號")
                .Columns.Add("淨重和")
                .Columns.Add("米數和")
                .Columns.Add("總價和")
                .Columns.Add("車次計數")
            End With

            For Each monthGroup In groupData.GroupBy(Function(x) x.Month).OrderBy(Function(x) x.Key)
                '寫入月份
                'cells(rowIndex, 1) = monthGroup.Key & "月"
                'BottomLine_Cell(cells(rowIndex, 1))
                'rowIndex += 1
                table.Rows.Add(monthGroup.Key & "月")

                For Each customerGroup In monthGroup.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                    For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                        For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)

                            'cells(rowIndex, 1) = customerGroup.Key
                            'cells(rowIndex, 2) = productGroup.Key
                            'cells(rowIndex, 3) = CarNoGroup.Key
                            '寫入          客戶               產品              車號
                            table.Rows.Add(customerGroup.Key, productGroup.Key, CarNoGroup.Key)

                            'BottomLine_Cell(rowIndex, 1, 3)

                            'rowIndex += 1

                            For Each carNoItem In CarNoGroup
                                For Each record In carNoItem.Data
                                    ' 寫入每條記錄的詳細資料
                                    'cells(rowIndex, 4) = record.Weight
                                    'cells(rowIndex, 5) = record.Meters
                                    'cells(rowIndex, 6) = record.Price
                                    table.Rows.Add("", "", "", record.Weight, record.Meters, record.Price)
                                    'rowIndex += 1
                                Next
                            Next

                            ' 寫入小計
                            'cells(rowIndex, 3) = "(小計)"
                            'cells(rowIndex, 4) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Weight)))
                            'cells(rowIndex, 5) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meters)))
                            'cells(rowIndex, 6) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Price)))
                            'cells(rowIndex, 7) = CarNoGroup.Count
                            table.Rows.Add("", "", "(小計)", CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Weight))), CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meters))) _
                                       , CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Price))), CarNoGroup.Count)

                            'TopLine_Cell(rowIndex, 3, 7)
                            ' 記住剛剛添加的「小計」行的行號
                            Dim subtotalRowIndex As Integer = table.Rows.Count

                            ' 添加上邊框到「小計」所在的行
                            Dim topBorderRange As Range = ws.Range(ws.Cells(subtotalRowIndex + 2, 1), ws.Cells(subtotalRowIndex + 2, table.Columns.Count))
                            topBorderRange.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            topBorderRange.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThin
                            'rowIndex += 2
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
        Public Sub GenerateMonthlyStatement(dtp As DateTimePicker, inOut As String)
            '撈資料
            Dim dtpValue = dtp.Value
            Dim startDate As New Date(dtpValue.Year, dtpValue.Month, 1)
            Dim endDate = startDate.AddMonths(1).AddDays(-1)

            Dim dic As New Dictionary(Of String, Object) From {
                {"startDate", startDate.ToString("yyyy/MM/dd")},
                {"endDate", endDate.ToString("yyyy/MM/dd")},
                {"inOut", inOut}
            }

            Dim dt = SelectTable(
                "SELECT 過磅日期, [客戶/廠商], 磅單序號, 空重, 總重, 單價, 每米噸數, 淨重, 米數, 總價, 產品名稱, 車牌號碼 FROM 過磅資料表 " &
                "WHERE 過磅日期 BETWEEN @endDate AND @startDate " &
                $"AND [進/出] = @inOut",
                dic
            )

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
                        .Date = row("過磅日期"),
                        .ID = row("磅單序號"),
                        .Empty = Math.Round(row("空重"), 3),
                        .Weight = Math.Round(row("總重"), 3),
                        .UnitPrice = row("單價"),
                        .TPM = row("每米噸數"),
                        .NetWeight = Math.Round(row("淨重"), 3),
                        .Meter = Math.Round(row("米數"), 3),
                        .Price = Math.Round(row("總價"), 3)
                    }).ToList
                }).ToList

            Dim rowIndex = 3

            For Each customerGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                '寫入客戶
                cells(rowIndex, 1) = "客戶:" & customerGroup.Key
                BottomLine_Cell(rowIndex, 1, 1)
                rowIndex += 1

                For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                    '寫入產品
                    cells(rowIndex, 1) = "產品:" & productGroup.Key
                    BottomLine_Cell(rowIndex, 1, 1)
                    rowIndex += 1

                    For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                        '寫入車號
                        cells(rowIndex, 1) = "車號:" & CarNoGroup.Key
                        BottomLine_Cell(rowIndex, 1, 1)
                        rowIndex += 1

                        For Each carNoItem In CarNoGroup
                            For Each record In carNoItem.Data
                                '寫入每條記錄的詳細資料
                                cells(rowIndex, 1) = record.Date
                                cells(rowIndex, 2) = record.ID
                                cells(rowIndex, 3) = record.Empty
                                cells(rowIndex, 4) = record.Weight
                                cells(rowIndex, 5) = record.UnitPrice
                                cells(rowIndex, 6) = record.TPM
                                cells(rowIndex, 7) = record.NetWeight
                                cells(rowIndex, 8) = record.Meter
                                cells(rowIndex, 9) = record.Price
                                rowIndex += 1
                            Next
                        Next

                        ' 寫入小計
                        cells(rowIndex, 6) = "(小計)"
                        cells(rowIndex, 7) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                        cells(rowIndex, 8) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))
                        cells(rowIndex, 9) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Price)))

                        TopLine_Cell(rowIndex, 6, 9)

                        rowIndex += 2
                    Next
                Next
            Next
        End Sub

        ''' <summary>
        ''' 日度對帳單
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyStatement(dtp As DateTimePicker, inOut As String)
            '撈資料
            Dim dateSelect As String = dtp.Value.ToString("yyyy/MM/dd")

            Dim dic As New Dictionary(Of String, Object) From {
                {"date", dateSelect},
                {"inOut", inOut}
            }

            Dim dt = SelectTable(
                "SELECT 過磅日期, [客戶/廠商], 磅單序號, 空重, 總重, 單價, 每米噸數, 淨重, 米數, 總價, 產品名稱, 車牌號碼 FROM 過磅資料表 " &
                $"WHERE 過磅日期 = @date " &
                $"AND [進/出] = @inOut",
                dic
            )

            '標題
            cells(1, 1) = $"{dateSelect} {inOut} 對帳單"

            ' 撈取當月客戶購買的產品所使用的車的資料
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

            Dim rowIndex = 3

            For Each customerGroup In datas.GroupBy(Function(x) x.Customer).OrderBy(Function(x) x.Key)
                '寫入客戶
                cells(rowIndex, 1) = "客戶:" & customerGroup.Key
                BottomLine_Cell(rowIndex, 1, 1)
                rowIndex += 1

                For Each productGroup In customerGroup.GroupBy(Function(x) x.Product).OrderBy(Function(x) x.Key)
                    '寫入產品
                    cells(rowIndex, 1) = "產品:" & productGroup.Key
                    BottomLine_Cell(rowIndex, 1, 1)
                    rowIndex += 1

                    For Each CarNoGroup In customerGroup.GroupBy(Function(x) x.CarNo).OrderBy(Function(x) x.Key)
                        '寫入車號
                        cells(rowIndex, 1) = "車號:" & CarNoGroup.Key
                        BottomLine_Cell(rowIndex, 1, 1)
                        rowIndex += 1

                        For Each carNoItem In CarNoGroup
                            For Each record In carNoItem.Data
                                '寫入每條記錄的詳細資料
                                cells(rowIndex, 1) = record.Date
                                cells(rowIndex, 2) = record.ID
                                cells(rowIndex, 3) = record.Empty
                                cells(rowIndex, 4) = record.Weight
                                cells(rowIndex, 5) = record.UnitPrice
                                cells(rowIndex, 6) = record.TPM
                                cells(rowIndex, 7) = record.NetWeight
                                cells(rowIndex, 8) = record.Meter
                                cells(rowIndex, 9) = record.Price
                                rowIndex += 1
                            Next
                        Next

                        ' 寫入小計
                        cells(rowIndex, 6) = "(小計)"
                        cells(rowIndex, 7) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.NetWeight)))
                        cells(rowIndex, 8) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Meter)))
                        cells(rowIndex, 9) = CarNoGroup.Sum(Function(x) x.Data.Sum(Function(d) CDbl(d.Price)))

                        TopLine_Cell(rowIndex, 6, 9)

                        rowIndex += 2
                    Next
                Next
            Next
        End Sub

        ''' <summary>
        ''' 月統計表
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyReport(dtp As DateTimePicker, inOut As String)
            '撈資料
            Dim dtpStart = dtp.Value
            Dim year As Integer = dtpStart.Year
            Dim month As Integer = dtpStart.Month

            '標題
            cells(1, 1) = $"{year} 年 {month} 月 {inOut} 統計表"

            '抓出當月所有客戶
            Dim dtCus = SelectTable(
                "SELECT DISTINCT [客戶/廠商] FROM 過磅資料表 " &
               $"WHERE YEAR(CDate(過磅日期)) = '{year}' " &
               $"AND MONTH(CDate(過磅日期)) = '{month}' " &
               $"AND [進/出] = '{inOut}'"
                )

            Dim rowIndex = 3

            For Each cus As DataRow In dtCus.Rows
                '列出客戶
                cells(rowIndex, 1) = "客戶:" & cus("客戶/廠商")
                BottomLine_Cell(cells(rowIndex, 1))
                rowIndex += 1

                '抓出所有車號
                Dim dtCarNo = SelectTable(
                    "SELECT DISTINCT 車牌號碼 FROM 過磅資料表 " &
                    $"WHERE YEAR(CDate(過磅日期)) = '{year}' " &
                    $"AND MONTH(CDate(過磅日期)) = '{month}' " &
                    $"AND [進/出] = '{inOut}' " &
                    $"AND [客戶/廠商] = '{cus("客戶/廠商")}'"
                )

                For Each carNo In dtCarNo.Rows
                    '列出車號
                    cells(rowIndex, 1) = "車號:" & carNo("車牌號碼")
                    BottomLine_Cell(cells(rowIndex, 1))
                    rowIndex += 1

                    '抓出產品
                    Dim dt = SelectTable(
                        "SELECT DISTINCT 產品名稱 FROM 過磅資料表 " &
                        $"WHERE YEAR(CDate(過磅日期)) = '{year}' " &
                        $"AND MONTH(CDate(過磅日期)) = '{month}' " &
                        $"AND [進/出] = '{inOut}' " &
                        $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                        $"AND 車牌號碼 = '{carNo("車牌號碼")}'"
                    )

                    Dim sumCarCount As Integer = 0
                    Dim sumWeight As Double = 0
                    Dim sumMeter As Double = 0
                    Dim sumPrice As Double = 0

                    For Each product In dt.Rows
                        '抓出車次、總出貨量(噸、米)、總金額
                        dt = SelectTable(
                            "SELECT SUM(淨重) AS total_weight, SUM(米數) AS total_meter, SUM(總價) AS total_price FROM 過磅資料表 " &
                            $"WHERE YEAR(CDate(過磅日期)) = '{year}' " &
                            $"AND MONTH(CDate(過磅日期)) = '{month}' " &
                            $"AND [進/出] = '{inOut}' " &
                            $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                            $"AND 車牌號碼 = '{carNo("車牌號碼")}'" &
                            $"AND 產品名稱 = '{product("產品名稱")}' "
                        )

                        '列出產品
                        cells(rowIndex, 1) = product("產品名稱")
                        '車次
                        Dim carCout = dt.Rows.Count
                        cells(rowIndex, 2) = carCout
                        '總出貨量(噸)
                        Dim weight = dt.Rows(0)("total_weight")
                        cells(rowIndex, 3) = Math.Round(weight, 3)
                        '總出貨量(米)
                        Dim meter = dt.Rows(0)("total_meter")
                        cells(rowIndex, 4) = Math.Round(meter, 3)
                        '總金額
                        Dim price = dt.Rows(0)("total_price")
                        cells(rowIndex, 5) = Math.Round(price, 3)
                        '列出車次

                        sumCarCount += carCout
                        sumWeight += weight
                        sumMeter += meter
                        sumPrice += price

                        rowIndex += 1
                    Next

                    For i As Integer = 1 To 5
                        TopLine_Cell(cells(rowIndex, i))
                    Next

                    cells(rowIndex, 1) = "(小計)"
                    cells(rowIndex, 2) = Math.Round(sumCarCount, 3)
                    cells(rowIndex, 3) = Math.Round(sumWeight, 3)
                    cells(rowIndex, 4) = Math.Round(sumMeter, 3)
                    cells(rowIndex, 5) = Math.Round(sumPrice, 3)

                    rowIndex += 2
                Next
            Next
        End Sub

        ''' <summary>
        ''' 月產品統計表
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateMonthlyProductStats(dtp As DateTimePicker, inOut As String)
            '撈資料
            Dim dtpStart = dtp.Value
            Dim year As Integer = dtpStart.Year
            Dim month As Integer = dtpStart.Month

            '標題
            cells(1, 1) = $"{year} 年 {month} 月 產品 {inOut} 統計表"

            '抓出當月產品
            Dim dt = SelectTable(
                "SELECT 產品名稱, COUNT(*) AS 車次, SUM(總重) AS 總重, SUM(米數) AS 米, SUM(總價) AS 總價 FROM 過磅資料表 " &
               $"WHERE YEAR(過磅日期) = {year} " &
               $"AND MONTH(過磅日期) = {month} " &
               $"AND [進/出] = '{inOut}' " &
                "GROUP BY 產品名稱 " &
                "ORDER BY 產品名稱"
                )

            Dim rowIndex = 3

            Dim sumCarCount As Integer = 0
            Dim sumWeight As Double = 0
            Dim sumMeter As Double = 0
            Dim sumPrice As Double = 0

            For Each row As DataRow In dt.Rows
                cells(rowIndex, 1) = row("產品名稱")
                cells(rowIndex, 2) = row("車次")
                sumCarCount += row("車次")
                cells(rowIndex, 3) = Math.Round(row("總重"), 3)
                sumWeight += row("總重")
                cells(rowIndex, 4) = Math.Round(row("米"), 3)
                sumMeter += row("米")
                cells(rowIndex, 5) = Math.Round(row("總價"), 3)
                sumPrice += row("總價")

                rowIndex += 1
            Next

            For i As Integer = 1 To 5
                TopLine_Cell(cells(rowIndex, i))
            Next

            cells(rowIndex, 1) = "(總計)"
            cells(rowIndex, 2) = Math.Round(sumCarCount, 3)
            cells(rowIndex, 3) = Math.Round(sumWeight, 3)
            cells(rowIndex, 4) = Math.Round(sumMeter, 3)
            cells(rowIndex, 5) = Math.Round(sumPrice, 3)
        End Sub

        ''' <summary>
        ''' 日產品統計表
        ''' </summary>
        ''' <param name="dtp"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyProductStats(dtp As DateTimePicker, inOut As String)
            '撈資料
            Dim d = dtp.Value

            '標題
            cells(1, 1) = $"{d.Year} 年 {d.Month} 月 {d.Day} 日 產品 {inOut} 統計表"

            '抓出當日產品
            Dim dt = SelectTable(
                "SELECT 產品名稱, COUNT(*) AS 車次, SUM(總重) AS 總重, SUM(米數) AS 米, SUM(總價) AS 總價 FROM 過磅資料表 " &
               $"WHERE DATEVALUE(過磅日期) = #{d.Date:yyyy/MM/dd}# " &
               $"AND [進/出] = '{inOut}' " &
                "GROUP BY 產品名稱 " &
                "ORDER BY 產品名稱"
                )

            Dim rowIndex = 3

            Dim sumCarCount As Integer = 0
            Dim sumWeight As Double = 0
            Dim sumMeter As Double = 0
            Dim sumPrice As Double = 0

            For Each row As DataRow In dt.Rows
                Dim weight As Double = Math.Round(row("總重"), 3)
                Dim meter As Double = Math.Round(row("米"), 3)
                Dim price As Double = Math.Round(row("總價"), 3)

                cells(rowIndex, 1) = row("產品名稱")
                cells(rowIndex, 2) = row("車次")
                sumCarCount += row("車次")
                cells(rowIndex, 3) = weight
                sumWeight += weight
                cells(rowIndex, 4) = meter
                sumMeter += meter
                cells(rowIndex, 5) = price
                sumPrice += price

                rowIndex += 1
            Next

            For i As Integer = 1 To 5
                TopLine_Cell(cells(rowIndex, i))
            Next

            cells(rowIndex, 1) = "(總計)"
            cells(rowIndex, 2) = sumCarCount
            cells(rowIndex, 3) = sumWeight
            cells(rowIndex, 4) = sumMeter
            cells(rowIndex, 5) = sumPrice

            rowIndex += 2
        End Sub

        ''' <summary>
        ''' 日產品客戶統計表
        ''' </summary>
        ''' <param name="dtpStart"></param>
        ''' <param name="dtpEnd"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyProductCustomerStats(dtpStart As DateTimePicker, dtpEnd As DateTimePicker, inOut As String)
            '撈資料
            Dim dStart = dtpStart.Value
            Dim dEnd = dtpEnd.Value
            Dim person As String = ""

            Select Case inOut
                Case "進貨"
                    person = "廠商"
                Case "出貨"
                    person = "客戶"
                Case Else
                    Exit Select
            End Select

            '標題
            cells(1, 1) = $"{dStart:yyyy/MM/dd} ~ {dEnd:yyyy/MM/dd} 產品{person} {inOut} 統計表"

            '抓出當日產品
            Dim dtCus = SelectTable(
                "SELECT DISTINCT [客戶/廠商] FROM 過磅資料表 " &
               $"WHERE DATEVALUE(過磅日期) >= #{dStart.Date:yyyy/MM/dd}# " &
               $"AND DATEVALUE(過磅日期) <= #{dEnd.Date:yyyy/MM/dd}# " &
               $"AND [進/出] = '{inOut}' "
                )

            Dim rowIndex = 3

            For Each cus As DataRow In dtCus.Rows
                '列出客戶
                cells(rowIndex, 1) = $"{person}:{cus("客戶/廠商")}"
                BottomLine_Cell(cells(rowIndex, 1))
                rowIndex += 1

                '抓出日期
                Dim dtDate = SelectTable(
                    "SELECT DISTINCT 過磅日期 FROM 過磅資料表 " &
                   $"WHERE DATEVALUE(過磅日期) >= #{dStart.Date:yyyy/MM/dd}# " &
                   $"AND DATEVALUE(過磅日期) <= #{dEnd.Date:yyyy/MM/dd}# " &
                   $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                   $"AND [進/出] = '{inOut}' "
                    )

                For Each d As DataRow In dtDate.Rows
                    '列出日期
                    cells(rowIndex, 1) = $"日期:{d("過磅日期")}"
                    BottomLine_Cell(cells(rowIndex, 1))
                    rowIndex += 1

                    Dim dt = SelectTable(
                        "SELECT 產品名稱, SUM(淨重) AS 總淨重, SUM(米數) AS 總米數, SUM(總價) AS 總金額 FROM 過磅資料表 " &
                        "WHERE DATEVALUE(過磅日期) = #" & CDate(d("過磅日期")).ToString("yyyy/MM/dd") & "# " &
                       $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                       $"AND [進/出] = '{inOut}' " &
                       "GROUP BY 產品名稱 " &
                       "ORDER BY 產品名稱"
                        )

                    Dim sumWeight As Double = 0
                    Dim sumMeter As Double = 0
                    Dim sumPrice As Double = 0

                    For Each row As DataRow In dt.Rows
                        Dim weight As Double = Math.Round(row("總淨重"), 3)
                        Dim meter As Double = Math.Round(row("總米數"), 3)
                        Dim price As Double = Math.Round(row("總金額"), 3)

                        cells(rowIndex, 1) = row("產品名稱")
                        cells(rowIndex, 2) = weight
                        sumWeight += weight
                        cells(rowIndex, 3) = meter
                        sumMeter += meter
                        cells(rowIndex, 4) = price
                        sumPrice += price

                        rowIndex += 1
                    Next

                    For i As Integer = 1 To 4
                        TopLine_Cell(cells(rowIndex, i))
                    Next

                    cells(rowIndex, 1) = "(總計)"
                    cells(rowIndex, 2) = sumWeight
                    cells(rowIndex, 3) = sumMeter
                    cells(rowIndex, 4) = sumPrice

                    rowIndex += 2
                Next
            Next
        End Sub

        ''' <summary>
        ''' 日客戶產品統計表
        ''' </summary>
        ''' <param name="dtpStart"></param>
        ''' <param name="dtpEnd"></param>
        ''' <param name="inOut"></param>
        Public Sub GenerateDailyCustomerProductStats(dtpStart As DateTimePicker, dtpEnd As DateTimePicker, inOut As String)
            '撈資料
            Dim dStart = dtpStart.Value
            Dim dEnd = dtpEnd.Value

            Dim person As String = ""

            Select Case inOut
                Case "進貨"
                    person = "廠商"
                Case "出貨"
                    person = "客戶"
                Case Else
                    Exit Select
            End Select

            '標題
            cells(1, 1) = $"{dStart:yyyy/MM/dd} ~ {dEnd:yyyy/MM/dd} {person}產品 {inOut} 統計表"

            '抓出區間內的客戶
            Dim dtCus = SelectTable(
                "SELECT DISTINCT [客戶/廠商] FROM 過磅資料表 " &
               $"WHERE DATEVALUE(過磅日期) >= #{dStart.Date:yyyy/MM/dd}# " &
               $"AND DATEVALUE(過磅日期) <= #{dEnd.Date:yyyy/MM/dd}# " &
               $"AND [進/出] = '{inOut}' "
                )

            Dim rowIndex = 3

            For Each cus As DataRow In dtCus.Rows
                '列出客戶
                cells(rowIndex, 1) = $"{person}:{cus("客戶/廠商")}"
                BottomLine_Cell(cells(rowIndex, 1))
                rowIndex += 1

                '抓出日期
                Dim dtDate = SelectTable(
                    "SELECT DISTINCT 過磅日期 FROM 過磅資料表 " &
                   $"WHERE DATEVALUE(過磅日期) >= #{dStart.Date:yyyy/MM/dd}# " &
                   $"AND DATEVALUE(過磅日期) <= #{dEnd.Date:yyyy/MM/dd}# " &
                   $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                   $"AND [進/出] = '{inOut}' "
                    )

                For Each d As DataRow In dtDate.Rows
                    '列出日期
                    cells(rowIndex, 1) = $"日期:{d("過磅日期")}"
                    BottomLine_Cell(cells(rowIndex, 1))
                    rowIndex += 1

                    Dim dt = SelectTable(
                        "SELECT 產品名稱, COUNT(*) AS 車次, SUM(淨重) AS 總淨重, SUM(米數) AS 總米數, SUM(總價) AS 總金額 FROM 過磅資料表 " &
                        "WHERE DATEVALUE(過磅日期) = #" & CDate(d("過磅日期")).ToString("yyyy/MM/dd") & "# " &
                       $"AND [客戶/廠商] = '{cus("客戶/廠商")}' " &
                       $"AND [進/出] = '{inOut}' " &
                       "GROUP BY 產品名稱 " &
                       "ORDER BY 產品名稱"
                        )

                    Dim sumCarCount As Integer = 0
                    Dim sumWeight As Double = 0
                    Dim sumMeter As Double = 0
                    Dim sumPrice As Double = 0

                    For Each row As DataRow In dt.Rows
                        Dim weight As Double = Math.Round(row("總淨重"), 3)
                        Dim meter As Double = Math.Round(row("總米數"), 3)
                        Dim price As Double = Math.Round(row("總金額"), 3)

                        cells(rowIndex, 1) = row("產品名稱")
                        cells(rowIndex, 2) = row("車次")
                        sumCarCount += row("車次")
                        cells(rowIndex, 3) = weight
                        sumWeight += weight
                        cells(rowIndex, 4) = meter
                        sumMeter += meter
                        cells(rowIndex, 5) = price
                        sumPrice += price

                        rowIndex += 1
                    Next

                    For i As Integer = 1 To 5
                        TopLine_Cell(cells(rowIndex, i))
                    Next

                    cells(rowIndex, 1) = "(總計)"
                    cells(rowIndex, 2) = sumCarCount
                    cells(rowIndex, 3) = sumWeight
                    cells(rowIndex, 4) = sumMeter
                    cells(rowIndex, 5) = sumPrice

                    rowIndex += 2
                Next
            Next
        End Sub

        ''' <summary>
        ''' 過磅單日統計表
        ''' </summary>
        ''' <param name="dtpStart"></param>
        ''' <param name="dtpEnd"></param>
        Public Sub GenerateWeighingDailyReport(dtpStart As DateTimePicker, dtpEnd As DateTimePicker)
            '撈資料
            Dim dStart = dtpStart.Value
            Dim dEnd = dtpEnd.Value

            '標題
            cells(1, 1) = $"{dStart:yyyy/MM/dd} ~ {dEnd:yyyy/MM/dd} 過磅單統計表"

            '抓出區間內的資料
            Dim dt = SelectTable(
                "SELECT 磅單序號, [進/出], [客戶/廠商], 車牌號碼, 產品名稱, 總重, 空重, 淨重, 米數 FROM 過磅資料表 " &
               $"WHERE DATEVALUE(過磅日期) >= #{dStart.Date:yyyy/MM/dd}# " &
               $"AND DATEVALUE(過磅日期) <= #{dEnd.Date:yyyy/MM/dd}# " &
               "ORDER BY 磅單序號"
                )

            Dim rowIndex = 3

            Dim sumWeight As Double = 0
            Dim sumEmptyWeight As Double = 0
            Dim sumNetWeight As Double = 0
            Dim sumMeter As Double = 0

            For Each row As DataRow In dt.Rows
                Dim weight As Double = Math.Round(row("總重"), 3)
                Dim emptyWeight As Double = Math.Round(row("空重"), 3)
                Dim netWeight As Double = Math.Round(row("淨重"), 3)
                Dim meter As Double = Math.Round(row("米數"), 3)

                cells(rowIndex, 1) = row("磅單序號")
                cells(rowIndex, 2) = row("進/出")
                cells(rowIndex, 3) = row("客戶/廠商")
                cells(rowIndex, 4) = row("車牌號碼")
                cells(rowIndex, 5) = row("產品名稱")
                cells(rowIndex, 6) = weight
                sumWeight += weight
                cells(rowIndex, 7) = emptyWeight
                sumEmptyWeight += emptyWeight
                cells(rowIndex, 8) = netWeight
                sumNetWeight += netWeight
                cells(rowIndex, 9) = meter
                sumMeter += meter

                rowIndex += 1
            Next

            For i As Integer = 1 To 9
                TopLine_Cell(cells(rowIndex, i))
            Next

            cells(rowIndex, 5) = "(總計)"
            cells(rowIndex, 6) = sumWeight
            cells(rowIndex, 7) = sumEmptyWeight
            cells(rowIndex, 8) = sumNetWeight
            cells(rowIndex, 9) = sumMeter
        End Sub

        Public Sub CreateNewReport(sheetName As String)
            Dim orgWb As Workbook = exl.Workbooks.Open(templatePath)
            Dim orgWs As Worksheet = orgWb.Worksheets(sheetName)
            wb = exl.Workbooks.Add
            orgWs.Copy(wb.Sheets(1))
            orgWb.Close(False)
            wb.Sheets(2).Delete
            ws = wb.Worksheets(sheetName)
            cells = ws.Cells
        End Sub

        Public Sub SaveReport()
            ws.SaveAs(savePath)
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
        ''' <param name="cell">目標儲存格</param>
        Protected Sub BottomLine_Cell(row As Integer, colStart As Integer, colEnd As Integer)
            For i = colStart To colEnd
                cells(row, i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                cells(row, i).Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
            Next
        End Sub
        <Obsolete>
        Protected Sub BottomLine_Cell(cell As Range)
            cell.Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            cell.Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThin
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
    End Class
End Namespace