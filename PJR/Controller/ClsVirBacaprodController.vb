Imports MySql.Data.MySqlClient
Imports IDM.Fungsi

Public Class ClsVirBacaprodController
#Region "SO IC"
    Public Sub insert1230CBR_Virbacaprod()
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
                End If

            Scom.CommandText = "SELECT COUNT(1) FROM vir_bacaprod WHERE jenis='SOIC-CBR';"
            TraceLog("insert1230CBR_Virbacaprod-Q1: " & Scom.CommandText, TipeTracelog.Info)

            If Scom.ExecuteScalar = 0 Then
                Scom.CommandText = "INSERT IGNORE INTO VIR_BACAPROD(JENIS,FILTER,KET,PROGRAM) VALUES ('SOIC-CBR', '', 'MEMO No 1230-CPS-23 SOIC-CBR','WDCP');"
                TraceLog("insert1230CBR_Virbacaprod-Q2: " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            TraceLog("Error insert1230CBR_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error insert1230CBR_Virbacaprod")
        Finally
            Scon.Close()
        End Try
    End Sub

    Public Function get1230CBR_Virbacaprod() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim resultFilter As Boolean = False

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT FILTER FROM vir_bacaprod WHERE JENIS='SOIC-CBR';"
            TraceLog("get1230CBR_Virbacaprod-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = "ON" Then
                resultFilter = True
            Else
                resultFilter = False
            End If

        Catch ex As Exception
            TraceLog("Error get1230CBR_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error get1230CBR_Virbacaprod")
        Finally
            Scon.Close()
        End Try

        Return resultFilter

    End Function

    Public Function get1230Tabel_Virbacaprod() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim resultFilter As Boolean = False

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM vir_bacaprod WHERE JENIS='SB_format2023';"
            TraceLog("get1230Tabel_Virbacaprod-Q1: " & Scom.CommandText)

            If Not Scom.ExecuteScalar = 0 Then
                Scom.CommandText = "SELECT FILTER FROM vir_bacaprod WHERE JENIS='SB_format2023';"
                TraceLog("get1230Tabel_Virbacaprod-Q2: " & Scom.CommandText)

                If Scom.ExecuteScalar = "ON" Then
                    resultFilter = True
                Else
                    resultFilter = False
                End If
            End If

        Catch ex As Exception
            TraceLog("Error get1230Tabel_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error get1230Tabel_Virbacaprod")
        Finally
            Scon.Close()
        End Try

        Return resultFilter
    End Function

    Public Function get1230TTL3_Virbacaprod() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim resultFilter As Boolean

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM vir_bacaprod WHERE jenis='1230_TTL3';"
            TraceLog("get1230TTL3_Virbacaprod-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = 0 Then
                resultFilter = False
            Else
                Scom.CommandText = "SELECT filter FROM vir_bacaprod WHERE jenis='1230_TTL3';"
                TraceLog("get1230TTL3_Virbacaprod-Q2: " & Scom.CommandText)

                If Scom.ExecuteScalar = "ON" Then
                    resultFilter = True
                Else
                    resultFilter = False
                End If
            End If

        Catch ex As Exception
            TraceLog("Error get1230TTL3_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error get1230TTL3_Virbacaprod")
        Finally
            Scon.Close()
        End Try

        Return resultFilter

    End Function
#End Region

#Region "SO ED"
    Public Sub insert1314CBR_Virbacaprod()
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT COUNT(1) FROM vir_bacaprod WHERE jenis='SOED-CBR';"
            TraceLog("insert1314CBR_Virbacaprod-Q1: " & Scom.CommandText, TipeTracelog.Info)

            If Scom.ExecuteScalar = 0 Then
                Scom.CommandText = "INSERT IGNORE INTO VIR_BACAPROD(JENIS,FILTER,KET,PROGRAM) VALUES ('SOED-CBR', '', 'MEMO No 1314-CPS-23 SOED-CBR','WDCP');"
                TraceLog("insert1314CBR_Virbacaprod-Q2: " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            TraceLog("Error insert1314CBR_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error insert1314CBR_Virbacaprod")
        Finally
            Scon.Close()
        End Try
    End Sub

    Public Function get1314CBR_Virbacaprod() As Boolean
        Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Scom As New MySqlCommand("", Scon)
        Dim resultFilter As Boolean = False

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT FILTER FROM vir_bacaprod WHERE JENIS='SOED-CBR';"
            TraceLog("get1314CBR_Virbacaprod-Q1: " & Scom.CommandText)

            If Scom.ExecuteScalar = "ON" Then
                resultFilter = True
            Else
                resultFilter = False
            End If

        Catch ex As Exception
            TraceLog("Error get1314CBR_Virbacaprod: " & ex.ToString)
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.OkOnly, "Error get1314CBR_Virbacaprod")
        Finally
            Scon.Close()
        End Try

        Return resultFilter

    End Function
#End Region
End Class
