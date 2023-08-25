Imports System.Management
''' <summary>
''' 產生序號
''' </summary>
Module modAuthorized
    '索引表
    Private arrChar As String = "05KLMNOY34JP8FG9AQRSHI12VWXTBC67DEUZ"

    Public Function GetAuthCode(SerialNumber As String) As String

        Dim returnStr = ""
        '用第四碼決定要向右shift多少位元
        Dim ShiftNum = InStr(1, arrChar, Mid(SerialNumber, 4, 1))
        For i = 1 To Len(SerialNumber)
            Dim tempNum = (InStr(1, arrChar, Mid(SerialNumber, i, 1)) + ShiftNum * i) Mod 36
            returnStr &= Mid(arrChar, tempNum + 1, 1)
        Next
        Return returnStr
    End Function

    ''' <summary>
    ''' 產生序號
    ''' </summary>
    ''' <returns></returns>
    Public Function GetSerialNumber() As String
        '取 主機板的 第二碼起的三碼
        '取 HD的 第六碼起的三碼
        '組合成 10碼的字串，第四碼用上面三個的第一碼查表後，加總後取餘數再查表

        '取 CPU的 第五碼起的三碼
        Dim returnStrA = ""
        Dim returnStrB = ""
        Dim returnStrC = ""
        Dim objMOS As ManagementObjectSearcher
        Dim objMOC As ManagementObjectCollection
        Dim objMO As ManagementObject

        '取 CPU的 第五碼起的三碼
        objMOS = New ManagementObjectSearcher("Select * From Win32_Processor")
        objMOC = objMOS.Get

        For Each objMO In objMOC
            returnStrA = Mid(objMO("ProcessorID"), 5, 3)
            objMO.Dispose()
            Exit For
        Next

        objMOS.Dispose()

        '取 主機板的 第二碼起的三碼
        Dim mc As New ManagementClass("Win32_BaseBoard") ' 註1
        mc.Scope.Options.EnablePrivileges = True ' 註2
        Dim sno As String

        For Each mo As ManagementObject In mc.GetInstances() ' 註3
            sno = mo("SerialNumber") ' 註4
            returnStrB = Mid(sno, 2, 3)
            Exit For
        Next

        mc.Dispose() ' 註5

        '取 HD的 第六碼起的三碼
        Dim qry As String = "SELECT * FROM Win32_PhysicalMedia Where Tag = '\\\\.\\PHYSICALDRIVE0'" ' 註1
        Dim mos As New ManagementObjectSearcher(qry) ' 註2
        mos.Scope.Options.EnablePrivileges = True ' 註3         

        For Each mo As ManagementObject In mos.Get() ' 註4
            sno = mo("SerialNumber").ToString.Trim ' 註5
            returnStrC = Mid(sno, 6, 3)
        Next

        mos.Dispose() ' 註6

        '插一個號碼在CPU與MB之間,用上面三個的第一碼查詢在arrChar的第幾個位置，加總後取餘數再查表,組合成 10碼的字串，
        Dim chkNum = (InStr(1, arrChar, Mid(returnStrA, 1, 1)) + InStr(1, arrChar, Mid(returnStrB, 1, 1)) + InStr(1, arrChar, Mid(returnStrC, 1, 1))) Mod 36
        Return returnStrA + Mid(arrChar, chkNum, 1) + returnStrB + returnStrC
    End Function
End Module
