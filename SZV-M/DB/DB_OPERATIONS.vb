Public Module DB_OPERATIONS
    Enum CleanType
        ALL
        ONLY_DO
        ONLY_PO
        ONLY_DO_PO
        ONLY_MAN
        ONLY_PE
        ONLY_WPR
    End Enum

    Public Sub CleanUP(Typ As CleanType)
        Dim DB As New DB_ENGINE
        Select Case Typ
            Case CleanType.ALL
                DB.RunSimpleCmd("DELETE FROM MAN;DELETE FROM PE; DELETE FROM PAY_DO; DELETE FROM PAY_PO; DELETE FROM POPEN_DO; DELETE FROM POPEN_PO; DELETE FROM RECIP_DO; DELETE FROM RECIP_PO;DELETE FROM IND4; DELETE FROM WPR")
                SetDBParam("PE_MD5", "0")
                SetDBParam("MAN_MD5", "0")
                SetDBStatus(0)
            Case CleanType.ONLY_DO
                DB.RunSimpleCmd("DELETE FROM MAN;DELETE FROM PE; DELETE FROM PAY_DO; DELETE FROM POPEN_DO; DELETE FROM RECIP_DO;DELETE FROM IND4")
            Case CleanType.ONLY_PO
                DB.RunSimpleCmd("DELETE FROM PAY_PO; DELETE FROM POPEN_PO; DELETE FROM RECIP_PO;DELETE FROM IND4")
            Case CleanType.ONLY_DO_PO
                DB.RunSimpleCmd("DELETE FROM PAY_DO; DELETE FROM PAY_PO; DELETE FROM POPEN_DO; DELETE FROM POPEN_PO; DELETE FROM RECIP_DO; DELETE FROM RECIP_PO;DELETE FROM IND4")
                SetDBStatus(0)
            Case CleanType.ONLY_MAN
                SetDBParam("MAN_MD5", "0")
                DB.RunSimpleCmd("DELETE FROM MAN")
            Case CleanType.ONLY_PE
                SetDBParam("PE_MD5", "0")
                DB.RunSimpleCmd("DELETE FROM PE")
            Case CleanType.ONLY_WPR
                DB.RunSimpleCmd("DELETE FROM WPR")
        End Select
        DB.RunSimpleCmd("VACUUM")
    End Sub

    Public Function GetDBStatus() As Integer
        Dim response As Integer
        Dim DB As New DB_ENGINE
        Dim cmd As New SQLite.SQLiteCommand(DB.Open) With {.CommandText = "SELECT PARAM_VALUE FROM DB_SETTINGS WHERE PARAM_NAME = 'DB_STATE'"}
        response = cmd.ExecuteScalar
        Return response
    End Function

    Public Sub SetDBStatus(status As String)
        Dim DB As New DB_ENGINE
        Dim cmd As New SQLite.SQLiteCommand(DB.Open) With {.CommandText = "UPDATE DB_SETTINGS SET PARAM_VALUE =:status WHERE PARAM_NAME = 'DB_STATE'"}
        cmd.Parameters.Add("status", DbType.String).Value = status
        cmd.ExecuteNonQuery()
    End Sub

    Public Function GetDBParam(ParamName As String) As String
        Dim DB As New DB_ENGINE
        Dim cmd As New SQLite.SQLiteCommand(DB.Open) With {.CommandText = "Select PARAM_VALUE from DB_SETTINGS where PARAM_NAME = :pname"}
        cmd.Parameters.Add("pname", DbType.String).Value = ParamName
        Return cmd.ExecuteScalar()
    End Function

    Public Sub SetDBParam(ParamName As String, ParamValue As String)
        Dim DB As New DB_ENGINE
        Dim cmd As New SQLite.SQLiteCommand(DB.Open) With {.CommandText = "UPDATE DB_SETTINGS SET PARAM_VALUE = :pvalue WHERE PARAM_NAME = :pname"}
        cmd.Parameters.Add("pname", DbType.String).Value = ParamName
        cmd.Parameters.Add("pvalue", DbType.String).Value = ParamValue
        cmd.ExecuteNonQuery()
    End Sub

End Module
