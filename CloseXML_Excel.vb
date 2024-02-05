Imports ClosedXML.Excel

Public Class CloseXML_Excel
    Implements IDisposable

    Private workbook As XLWorkbook
    Private worksheet As IXLWorksheet

    Public Sub New(filePath As String)
        Try
            workbook = New XLWorkbook(filePath)
            worksheet = workbook.Worksheets.First()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 設定指定列的欄寬。
    ''' </summary>
    ''' <param name="columnIndex">列的索引，從 1 開始。</param>
    ''' <param name="width">欄寬值。</param>
    Public Sub SetColumnWidth(columnIndex As Integer, width As Double)
        worksheet.Column(columnIndex).Width = width
    End Sub

    ''' <summary>
    ''' 設定指定行的列高。
    ''' </summary>
    ''' <param name="rowIndex">行的索引，從 1 開始。</param>
    ''' <param name="height_cm">列高值。</param>
    Public Sub SetRowHeight(rowIndex As Integer, height_cm As Double)
        worksheet.Row(rowIndex).Height = CmToPoints(height_cm)
    End Sub

    Public Sub SelectWorksheet(sheetName As String)
        Try
            worksheet = workbook.Worksheet(sheetName)
        Catch ex As Exception
            Console.WriteLine($"Sheet '{sheetName}' not found: {ex.Message}")
        End Try
    End Sub

    Public Sub WriteToCell(rowIndex As Integer, columnIndex As Integer, content As String, Optional formatOptions As CellFormatOptions = Nothing)
        Dim cell = worksheet.Cell(rowIndex, columnIndex)
        cell.Value = content

        If formatOptions IsNot Nothing Then
            cell.Style.Font.FontName = formatOptions.FontName
            cell.Style.Font.FontSize = formatOptions.FontSize
            cell.Style.Font.Bold = formatOptions.IsBold
            cell.Style.Alignment.Horizontal = If(formatOptions.HorizontalCenter, XLAlignmentHorizontalValues.Center, XLAlignmentHorizontalValues.Left)
            cell.Style.Alignment.Vertical = If(formatOptions.VerticalCenter, XLAlignmentVerticalValues.Center, XLAlignmentVerticalValues.Top)
            cell.Style.Alignment.WrapText = formatOptions.WrapText
            If formatOptions.VerticalText Then cell.Style.Alignment.TextRotation = 255
        End If
    End Sub

    Public Function SaveAs(fileName As String) As Boolean
        Dim newFileName = fileName & ".xlsx"

        Try
            Dim sfd = New SaveFileDialog With {
                .Filter = "Excel文件|*.xlsx",
                .FileName = newFileName
            }

            If sfd.ShowDialog = DialogResult.OK Then
                workbook.SaveAs(sfd.FileName)
                MsgBox("報表建立成功!")
                Return True
            End If

        Catch ex As Exception
            MsgBox($"Error saving file '{newFileName}': {ex.Message}")
        End Try

        Return False
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        workbook?.Dispose()
    End Sub

    Public Class CellFormatOptions
        Public Property FontName As String = "新細明體"
        Public Property FontSize As Double = 11
        Public Property IsBold As Boolean = False
        ''' <summary>
        ''' 水平置中
        ''' </summary>
        ''' <returns></returns>
        Public Property HorizontalCenter As Boolean = False
        ''' <summary>
        ''' 垂直置中
        ''' </summary>
        ''' <returns></returns>
        Public Property VerticalCenter As Boolean = False
        ''' <summary>
        ''' 自動換行
        ''' </summary>
        ''' <returns></returns>
        Public Property WrapText As Boolean = False
        ''' <summary>
        ''' 垂直文字
        ''' </summary>
        ''' <returns></returns>
        Public Property VerticalText As Boolean = False
    End Class

    Private Function CmToPoints(cm As Double) As Double
        Return cm * 28.35
    End Function
End Class
