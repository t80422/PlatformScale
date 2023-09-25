Imports PlatformScale.Models

Public Class ReportService
    Private Const reportQuery As String = "SELECT 過磅日期, [客戶/廠商], 磅單序號, 空重, 總重, 單價, 每米噸數, 淨重, 米數, 總價, 產品名稱, 車牌號碼 FROM 過磅資料表 "

    ''' <summary>
    ''' 取得年度報表數據
    ''' </summary>
    ''' <param name="year">指定年分</param>
    ''' <param name="inOut">進出貨</param>
    ''' <returns>報表數據</returns>
    Public Function GetYearlyData(year As Integer, inOut As String) As List(Of ReportData)
        ValidateYear(year)
        ValidateInOut(inOut)

        Dim sql As String = "WHERE YEAR(CDATE(過磅日期)) = @year AND [進/出] = @inOut"
        Dim dic As New Dictionary(Of String, Object) From {
            {"@year", year},
            {"@inOut", inOut}
        }
        Return FetchData(sql, dic)
    End Function

    ''' <summary>
    ''' 取得月度報表數據
    ''' </summary>
    ''' <param name="year">指定年分</param>
    ''' <param name="month">指定月份</param>
    ''' <param name="inOut">進出貨</param>
    ''' <returns>報表數據</returns>
    Public Function GetMonthlyData(year As Integer, month As Integer, inOut As String) As List(Of ReportData)
        ValidateYear(year)
        ValidateMonth(month)
        ValidateInOut(inOut)

        Dim sql As String = "WHERE YEAR(CDATE(過磅日期)) = @year AND MONTH(CDATE(過磅日期)) = @month AND [進/出] = @inOut"
        Dim dic As New Dictionary(Of String, Object) From {
            {"@year", year},
            {"@month", month},
            {"@inOut", inOut}
        }
        Return FetchData(sql, dic)
    End Function

    ''' <summary>
    ''' 取得日度報表數據
    ''' </summary>
    ''' <param name="d">指定日期</param>
    ''' <param name="inOut">進出貨</param>
    ''' <returns>報表數據</returns>
    Public Function GetDailyData(d As Date, inOut As String) As List(Of ReportData)
        ValidateInOut(inOut)

        Dim dateSelect As String = d.ToString("yyyy/MM/dd")
        Dim sql As String = "WHERE 過磅日期 = @dateSelect AND [進/出] = @inOut"
        Dim dic As New Dictionary(Of String, Object) From {
            {"@dateSelect", dateSelect},
            {"@inOut", inOut}
        }
        Return FetchData(sql, dic)
    End Function

    Private Function FetchData(whereCondition As String, parameters As Dictionary(Of String, Object)) As List(Of ReportData)
        Dim list As New List(Of ReportData)
        Try
            Dim dt As DataTable = SelectTable(reportQuery & whereCondition, parameters)

            For Each row As DataRow In dt.Rows
                Dim data As New ReportData With {
                    .WeighingDate = row("過磅日期"),
                    .ClientOrSupplier = row("客戶/廠商"),
                    .No = If(row.Table.Columns.Contains("磅單序號"), row("磅單序號"), Nothing),
                    .EmptyWeight = If(row.Table.Columns.Contains("空重"), row("空重"), Nothing),
                    .Weight = If(row.Table.Columns.Contains("總重"), row("總重"), Nothing),
                    .UnitPrice = If(row.Table.Columns.Contains("單價"), row("單價"), Nothing),
                    .TPM = If(row.Table.Columns.Contains("每米噸數"), row("每米噸數"), Nothing),
                    .NetWeight = row("淨重"),
                    .Meter = row("米數"),
                    .Price = row("總價"),
                    .Product = row("產品名稱"),
                    .CarNo = row("車牌號碼")
                }

                list.Add(data)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return list
    End Function

    Private Sub ValidateYear(year As Integer)
        Dim currentYear = Date.Now.Year
        If year < 1900 OrElse year > currentYear Then Throw New ArgumentException("無效的年")
    End Sub

    Private Sub ValidateMonth(month As Integer)
        If month < 1 OrElse month > 12 Then Throw New ArgumentException("無效的月")
    End Sub

    Private Sub ValidateInOut(inOut As String)
        Dim validValue = New List(Of String) From {"進貨", "出貨"}
        If Not validValue.Contains(inOut) Then Throw New ArgumentException("無效的 進/出")
    End Sub
End Class
