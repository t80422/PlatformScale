Public Interface ICarRep
    ''' <summary>
    ''' 刪除
    ''' </summary>
    ''' <param name="carNum">車號</param>
    ''' <param name="owner">車主</param>
    Sub Delete(carNum As String, owner As String)
End Interface
