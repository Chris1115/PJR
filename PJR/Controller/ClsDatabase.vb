Imports MySql.Data.MySqlClient

Public Class ClsDatabase
    Private utility As New Utility

    Public Function DropTableTracelogHanheld() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            mcom.CommandText = "DROP TABLE IF EXISTS `tracelog_handheld`; "
            mcom.ExecuteNonQuery()

        Catch ex As Exception
            result = False
            utility.TraceLogTxt("Error " & vbCrLf & "DropTableTracelogHanheld " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        Finally
            conn.Close()
        End Try

        Return result
    End Function

    Public Function BersihTableTraceLog() As Boolean
        'Dim connection As New ClsConnection
        Dim conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim result As Boolean = True
        Dim mcom As New MySqlCommand("", conn)
        Dim sqlQuery As String = ""

        Dim tmpDate As Date = Date.Today.AddMonths(-4) 'tanggal 4bulan yang lalu

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            sqlQuery = "DELETE FROM tracelog WHERE AppName LIKE 'HandheldIDM.exe%' AND TGL <'" & Format(tmpDate, "yyyy-MM-dd") & "'; "

            mcom.CommandText = sqlQuery
            mcom.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            utility.TraceLogTxt("Error " & vbCrLf & "BersihTableTraceLog" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Return False
        Finally
            conn.Close()
        End Try
    End Function

End Class
