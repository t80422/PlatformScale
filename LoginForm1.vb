Public Class LoginForm1
    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        Dim rows = SelectTable($"SELECT * FROM 密碼資料表 WHERE 名稱 = '{txtUser.Text}' AND 密碼 = '{txtPsw.Text}'").Rows
        If rows.Count > 0 Then
            Select Case rows(0)("權限")
                Case "1" '可操作一般過磅與新增資料, 但無法修改刪除及使用系統設定的功能
                    frmMain.tpSystem.Parent = Nothing
                    frmMain.grpAutoManu.Enabled = False
                    frmMain.permissions = 1
                Case "2" '可修改過磅與刪除過磅資料, 亦可手動與自動過磅, 但仍無法操作系統設定功能
                    frmMain.tpSystem.Parent = Nothing
                    frmMain.permissions = 2
                Case "3" '系統設定的最高等級, 可操作任何畫面與資料
                    If frmMain.tpSystem.Parent Is Nothing Then
                        frmMain.tpSystem.Parent = frmMain.tabMain
                        frmMain.tabMain.TabPages.Remove(frmMain.tpLogout)
                        frmMain.tabMain.TabPages.Add(frmMain.tpLogout)
                    End If
                    frmMain.permissions = 3
                Case Else
                    MsgBox("查無此權限")
                    Exit Sub
            End Select
            frmMain.Show()
            frmMain.rdoShipment.Checked = True
            frmMain.tabMain.SelectedIndex = 0
            Hide()
        Else
            MsgBox("帳號密碼錯誤")
        End If
        txtUser.Clear()
        txtPsw.Clear()
        txtUser.Focus()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class
