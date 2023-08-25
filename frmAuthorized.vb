Imports System.IO

Public Class frmAuthorized
    Private filePath = Path.Combine(Application.StartupPath, "Authorized.set")

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        If txtAuthorization.Text = GetAuthCode(txtSerialNum.Text) Then
            '儲存授權碼
            File.WriteAllText(filePath, txtAuthorization.Text)
            MsgBox("授權成功")
            LoginForm1.Show()
            Hide()
        Else
            MsgBox("授權碼錯誤", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub frmAuthorized_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Hide()
        Dim seriaNum = GetSerialNumber()

        If Not File.Exists(filePath) Then
            File.Create(filePath).Close()
        Else
            If File.ReadAllText(filePath) = GetAuthCode(seriaNum) Then
                LoginForm1.Show()
                Exit Sub
            End If
        End If

        txtSerialNum.Text = seriaNum
        Show()
    End Sub
End Class