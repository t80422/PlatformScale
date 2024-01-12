Imports System.Data.OleDb
Imports System.IO

Module modAccess
    Public conn As OleDbConnection
    Private title As String = "資料庫"
    Private dataSource As String
    Private passWord As String
    Private connStr As String

    '初始化
    Sub New()
        Try
            Dim dbSet = ReadConfigFile("DB.set")

            If dbSet Is Nothing Then
                SetDatabase()
            Else
                '檢查Db是否在位置上
                If Not File.Exists(dbSet(0)) Then SetDatabase()
                dataSource = dbSet(0)
                passWord = dbSet(1)
                connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dataSource};Jet OLEDB:Database Password={passWord}"
                conn = New OleDbConnection(connStr)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        TestConnect()
    End Sub

    Public Sub TestConnect()
        Try
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            SetDatabase()
        End Try
        conn.Close()
    End Sub

    Public Sub SetDatabase()
        dataSource = InputBox("請輸入資料庫路徑")
        If dataSource = "" Then End

        passWord = InputBox("請輸入密碼")

        Dim content = dataSource & vbCrLf & passWord
        CreateOrUpdateConfigFile("DB.set", content)
        connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dataSource};Jet OLEDB:Database Password={passWord}"
        conn = New OleDbConnection(connStr)
    End Sub

    ''' <summary>
    ''' 查詢資料表
    ''' </summary>
    ''' <param name="sSQL">完整的查詢語法</param>
    ''' <returns></returns>
    Public Function SelectTable(sSQL As String) As DataTable
        Dim dt As New DataTable()

        Try
            conn.Open()
            Using cmd As New OleDbCommand(sSQL, conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, title)
        End Try

        conn.Close()

        Return dt
    End Function

    Public Function SelectTable(query As String, parameters As Dictionary(Of String, Object)) As DataTable
        Dim dt As New DataTable()
        Try
            conn.Open()
            Using cmd As New OleDbCommand(query, conn)
                parameters.ToList.ForEach(Function(p) cmd.Parameters.AddWithValue(p.Key, p.Value))

                Dim adapter As New OleDbDataAdapter(cmd)

                adapter.Fill(dt)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, title)
        End Try
        conn.Close()
        Return dt
    End Function

    ''' <summary>
    ''' 檢查必填欄位
    ''' </summary>
    ''' <param name="container">父控制項</param>
    ''' <param name="required">必填清單</param>
    ''' <returns></returns>
    Public Function CheckText(container As Control, required As List(Of String)) As Boolean
        For Each txt In container.Controls.OfType(Of TextBox)().Where(Function(x) required.Contains(x.Tag.ToString) AndAlso String.IsNullOrWhiteSpace(x.Text))
            MsgBox(txt.Tag.ToString & " 不能空白")
            txt.Focus()
            Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' 新增資料到資料表(欄位名稱有特殊字元,不能使用Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value)))
    ''' </summary>
    ''' <param name="sTable">資料表名稱</param>
    ''' <param name="dicData">key:欄位名稱 Value:值</param>
    ''' <returns></returns>
    Public Function InserTable(sTable As String, dicData As Dictionary(Of String, String)) As Boolean
        Dim result As Boolean
        Dim cmd As New OleDbCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Values.Select(Function(x) $"'{Trim(x)}'"))})", conn)
        Try
            conn.Open()
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 新增資料到資料表
    ''' </summary>
    ''' <param name="sTable">資料表名稱</param>
    ''' <param name="dicData">key:欄位名稱 Value:值</param>
    ''' <returns></returns>
    Public Function InserTable(sTable As String, dicData As Dictionary(Of String, Object)) As Boolean
        Dim result As Boolean
        Dim cmd As New OleDbCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Keys.Select(Function(x) $"@{x}"))})", conn)
        Try
            conn.Open()
            For Each kvp In dicData
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 更新表格
    ''' </summary>
    ''' <param name="table">表格名稱</param>
    ''' <param name="dicFields">更新對象集合</param>
    ''' <param name="condition">條件 xxx=xxx</param>
    Public Function UpdateTable(table As String, dicFields As Dictionary(Of String, String), condition As String) As Boolean
        Dim result As Boolean
        Try
            conn.Open()
            Dim sql = $"UPDATE {table} SET "
            Dim lst As New List(Of String)
            For Each kvp In dicFields
                lst.Add($"{kvp.Key} = '{kvp.Value}'")
            Next
            sql += String.Join(",", lst) + $" WHERE {condition}"
            Dim cmd As New OleDbCommand(sql, conn)
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 表格刪除
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="sWhere">條件 xxx=xxx</param>
    ''' <returns></returns>
    Public Function DeleteTable(sTable As String, sWhere As String) As Boolean
        Dim rowsAffected As Integer
        Dim cmd As New OleDbCommand($"DELETE FROM {sTable} WHERE {sWhere}", conn)
        Try
            conn.Open()
            rowsAffected = cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return rowsAffected > 0
    End Function
    ''' <summary>
    ''' 表格刪除
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function DeleteTable(table As String, condition As Dictionary(Of String, String)) As Boolean
        Dim rowsAffected As Integer
        Dim list As New List(Of String)
        For Each kvp In condition
            list.Add($"{kvp.Key} = @{kvp.Key}")
        Next

        Dim cmd As New OleDbCommand($"DELETE FROM {table} WHERE {String.Join(" AND ", list)}", conn)
        Try
            conn.Open()
            For Each kvp In condition
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next
            rowsAffected = cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        conn.Close()
        Return rowsAffected > 0
    End Function

    ''' <summary>
    ''' 檢查是否重複新增
    ''' </summary>
    ''' <param name="selectFrom">SQL前半段</param>
    ''' <param name="list">條件,輸入控制項會自動取得Tag(欄位名稱),Text(值)</param>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    <Obsolete>
    Public Function CheckDuplication(selectFrom As String, list As List(Of Object), dgv As DataGridView) As Boolean
        Dim sql = selectFrom + $" WHERE {String.Join(" AND ", list.Select(Function(x) $"{x.tag} = '{x.text}'"))}"
        Dim dt = SelectTable(sql)
        If dt.Rows.Count > 0 Then
            MsgBox("重複資料")
            dgv.DataSource = dt
            Return False
        End If
        Return True
    End Function
    ''' <summary>
    ''' 檢查是否重複新增
    ''' </summary>
    ''' <param name="selectFrom">SQL前半段</param>
    ''' <param name="dic">條件,key:欄位 value:值</param>
    ''' <param name="dgv">欲顯示的DataGridView</param>
    ''' <returns></returns>
    Public Function CheckDuplication(selectFrom As String, dic As Dictionary(Of String, String), Optional dgv As DataGridView = Nothing) As Boolean
        '修正參數List>Dictionary,如果遇到DateTimePicker就會遇到可能取的值不是想要的
        Dim lst As List(Of String) = dic.Select(Function(kvp) $"{kvp.Key} = {kvp.Value}").ToList
        Dim sql = selectFrom + $" WHERE {String.Join(" AND ", lst)}"
        Dim dt = SelectTable(sql)
        If dt.Rows.Count > 0 Then
            MsgBox("重複資料")
            If dgv IsNot Nothing Then dgv.DataSource = dt
            Return False
        End If
        Return True
    End Function
End Module