Public Class LoginForm1
    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        Dim rows = SelectTable($"SELECT * FROM �K�X��ƪ� WHERE �W�� = '{txtUser.Text}' AND �K�X = '{txtPsw.Text}'").Rows
        If rows.Count > 0 Then
            Select Case rows(0)("�v��")
                Case "1" '�i�ާ@�@��L�S�P�s�W���, ���L�k�ק�R���ΨϥΨt�γ]�w���\��
                    frmMain.tpSystem.Parent = Nothing
                    frmMain.grpAutoManu.Enabled = False
                    frmMain.permissions = 1
                Case "2" '�i�ק�L�S�P�R���L�S���, ��i��ʻP�۰ʹL�S, �����L�k�ާ@�t�γ]�w�\��
                    frmMain.tpSystem.Parent = Nothing
                    frmMain.permissions = 2
                Case "3" '�t�γ]�w���̰�����, �i�ާ@����e���P���
                    If frmMain.tpSystem.Parent Is Nothing Then
                        frmMain.tpSystem.Parent = frmMain.tabMain
                        frmMain.tabMain.TabPages.Remove(frmMain.tpLogout)
                        frmMain.tabMain.TabPages.Add(frmMain.tpLogout)
                    End If
                    frmMain.permissions = 3
                Case Else
                    MsgBox("�d�L���v��")
                    Exit Sub
            End Select
            frmMain.Show()
            frmMain.rdoShipment.Checked = True
            frmMain.tabMain.SelectedIndex = 0
            Hide()
        Else
            MsgBox("�b���K�X���~")
        End If
        txtUser.Clear()
        txtPsw.Clear()
        txtUser.Focus()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class
