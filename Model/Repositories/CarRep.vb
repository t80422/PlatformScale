Imports System.Data.OleDb

Public Class CarRep
    Implements ICarRep

    Public Sub Delete(carNum As String, owner As String) Implements ICarRep.Delete
        Try
            conn.Open()

            Dim sql As String = "DELETE FROM 車籍資料表 WHERE 車號 = @carNum AND 車主 = @owner"

            Using command As New OleDbCommand(sql, conn)
                command.Parameters.AddWithValue("@carNum", carNum)
                command.Parameters.AddWithValue("@owner", owner)
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw New InvalidOperationException($"刪除車籍資料時發生錯誤：{ex.Message}", ex)
        End Try
    End Sub
End Class
