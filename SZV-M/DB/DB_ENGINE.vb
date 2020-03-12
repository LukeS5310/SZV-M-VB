Imports System.Data.SQLite

Public Class DB_ENGINE

#Region "Settings"
    ReadOnly ConnStr As String = "Data Source=" & Application.StartupPath & "\database.sqlite;"
    ReadOnly conn As New SQLiteConnection(ConnStr)
#End Region

#Region "Functions"
    Public Function Open() As SQLiteConnection
        Try
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return conn
    End Function
    Public Function Close() As SQLiteConnection
        Try
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return conn
    End Function

    Public Function RunSimpleCmd(cmdtext As String) As Boolean
        Try
            Dim cmd As New SQLiteCommand(Open())
            cmd.CommandText = cmdtext
            cmd.ExecuteNonQuery()
            Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Close()
            Return False
        End Try
    End Function

#End Region
End Class