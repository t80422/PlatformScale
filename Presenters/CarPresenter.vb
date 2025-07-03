Public Class CarPresenter
    Private ReadOnly _view As ICarView
    Private ReadOnly _repository As ICarRep
    Private currentData As Cars

    Public Sub New(view As ICarView, rep As ICarRep)
        _view = view
        _repository = rep
    End Sub

    Public Sub Delete(carNum As String, owner As String)
        Try
            _repository.Delete(carNum, owner)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
