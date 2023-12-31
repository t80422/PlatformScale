﻿Imports System.Text

Public Class frmScaleQuery
    Private Sub ScaleQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbCustomer.DataSource = SelectTable("SELECT 代號, 簡稱 FROM 客戶資料表")
        cmbCustomer.DisplayMember = "簡稱"
        cmbCustomer.ValueMember = "代號"
        cmbManufacturer.DataSource = SelectTable("SELECT 代號, 簡稱 FROM 廠商資料表")
        cmbManufacturer.DisplayMember = "簡稱"
        cmbManufacturer.ValueMember = "代號"
        cmbCarNo.DataSource = SelectTable("SELECT 車號 FROM 車籍資料表")
        cmbCarNo.DisplayMember = "車號"
        cmbCarNo.ValueMember = "車號"
        cmbProduct.DataSource = SelectTable("SELECT 代號, 品名 FROM 產品資料表")
        cmbProduct.DisplayMember = "品名"
        cmbProduct.ValueMember = "代號"
        GroupBox2.Controls.OfType(Of ComboBox).ToList.ForEach(Sub(cmb) cmb.SelectedIndex = -1)
    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        Dim mc As MonthCalendar = sender
        txtDate.Text = mc.SelectionStart.ToString("yyyy/MM/dd")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtDate.Clear()
        GroupBox2.Controls.OfType(Of ComboBox).ToList.ForEach(Sub(cmb)
                                                                  cmb.SelectedIndex = -1
                                                                  cmb.ResetText()
                                                              End Sub)
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Dim list As New List(Of String)
        For Each grp In Controls.OfType(Of GroupBox)
            grp.Controls.OfType(Of Control).Where(Function(ctrls) ctrls.GetType.Name <> "MonthCalendar" AndAlso Not String.IsNullOrEmpty(ctrls.Tag) AndAlso Not String.IsNullOrEmpty(ctrls.Text)).ToList.
                ForEach(Sub(x) list.Add($"{x.Tag} = '{x.Text}'"))
        Next
        frmMain.dgv過磅.DataSource = SelectTable($"SELECT * FROM 過磅資料表 WHERE {String.Join(" AND ", list)}")
        Close()
    End Sub

    '客戶 廠商只能二擇一
    Private Sub cmbCustomer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCustomer.SelectedIndexChanged
        If sender.SelectedIndex <> -1 Then
            cmbManufacturer.SelectedIndex = -1
            cmbManufacturer.ResetText()
        End If
    End Sub

    '客戶 廠商只能二擇一
    Private Sub cmbManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManufacturer.SelectedIndexChanged
        If sender.SelectedIndex <> -1 Then
            cmbCustomer.SelectedIndex = -1
            cmbCustomer.ResetText()
        End If
    End Sub
End Class