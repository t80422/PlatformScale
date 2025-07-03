Public Interface ICarView
    Event DeleteCar As EventHandler(Of DeleteRequestedEventArgs)
End Interface

Public Class DeleteRequestedEventArgs
    Inherits EventArgs

    Public Property CarNum As String
    Public Property Owner As String

    Public Sub New(carNum As String, owner As String)
        Me.CarNum = carNum
        Me.Owner = owner
    End Sub
End Class