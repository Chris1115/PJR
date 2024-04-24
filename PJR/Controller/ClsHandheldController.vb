Imports MySql.Data.MySqlClient
Imports IDM.Fungsi

Public Class ClsHandheldController
    Dim Scon As MySqlConnection = ClsConnection.GetConnection.Clone
    Dim Scom As New MySqlCommand("", Scon)

    Public Sub cekTabelHandheld()
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SHOW TABLES LIKE 'handheld_device';"
            TraceLog("cekTabelHandheld-Q1: " & Scom.CommandText)
            If ("" & Scom.ExecuteScalar) = "" Then
                Scom.CommandText = "CREATE TABLE `handheld_device`("
                Scom.CommandText &= " ip_address VARCHAR(50) DEFAULT '',"
                Scom.CommandText &= " socketID VARCHAR(50) DEFAULT '',"
                Scom.CommandText &= " jenis_so VARCHAR(50) DEFAULT '',"
                Scom.CommandText &= " lokasi_so VARCHAR(50) DEFAULT '',"
                Scom.CommandText &= " PRIMARY KEY(ip_address)"
                Scom.CommandText &= ");"
                TraceLog("cekTabelHandheld-Q2: " & Scom.CommandText)
            Else
                Scom.CommandText = "TRUNCATE TABLE handheld_device;"
                TraceLog("cekTabelHandheld-Q2: " & Scom.CommandText)
            End If

            Scom.ExecuteNonQuery()
        Catch ex As Exception
            TraceLog("Error cekTabelHandheld: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        End Try
    End Sub

    Public Sub addDevice(handheldModel As ClsHandheld)
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT * FROM handheld_device WHERE "
            Scom.CommandText &= "ip_address='" & handheldModel.ipAddress & "';"
            TraceLog("addDevice-Q1: " & Scom.CommandText)

            If ("" & Scom.ExecuteScalar) = "" Then
                Scom.CommandText = "INSERT INTO handheld_device (ip_address,socketID) "
                Scom.CommandText &= "VALUES ( '" & handheldModel.ipAddress & "', "
                Scom.CommandText &= "'" & handheldModel.socketID & "');"
                TraceLog("addDevice-Q2: " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            Else
                Scom.CommandText = "UPDATE handheld_device SET "
                Scom.CommandText &= "socketID='" & handheldModel.socketID & "' "
                Scom.CommandText &= "WHERE ip_address='" & handheldModel.ipAddress & "';"
                TraceLog("addDevice-Q2: " & Scom.CommandText)
                Scom.ExecuteNonQuery()
            End If

        Catch ex As Exception
            TraceLog("Error addDevice: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        End Try
    End Sub

    Public Sub addSO(handheldModel As ClsHandheld)
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "UPDATE handheld_device SET "
            Scom.CommandText &= "jenis_so='" & handheldModel.jenis_so & "', "
            Scom.CommandText &= "lokasi_so='" & handheldModel.lokasi_so & "', "
            Scom.CommandText &= "socketID='" & handheldModel.socketID & "' "
            Scom.CommandText &= "WHERE ip_address='" & handheldModel.ipAddress & "';"
            TraceLog("addSO-Q1: " & Scom.CommandText)
            Scom.ExecuteNonQuery()

        Catch ex As Exception
            TraceLog("Error addSO: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        End Try
    End Sub

    Public Function getLokasiSO(handheldModel As ClsHandheld) As String
        Dim lokasi As String = ""

        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "SELECT lokasi_so FROM handheld_device "
            Scom.CommandText &= "WHERE ip_address='" & handheldModel.ipAddress & "';"
            TraceLog("getLokasiSO-Q1: " & Scom.CommandText)
            lokasi = Scom.ExecuteScalar

        Catch ex As Exception
            TraceLog("Error getLokasiSO: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        End Try

        Return lokasi
    End Function

    Public Sub deleteDevice(handheldModel As ClsHandheld)
        Try
            If Scon.State = ConnectionState.Closed Then
                Scon.Open()
            End If

            Scom.CommandText = "DELETE FROM handheld_device "
            Scom.CommandText &= "WHERE ip_address='" & handheldModel.ipAddress & "';"
            TraceLog("deleteDevice-Q1: " & Scom.CommandText)

        Catch ex As Exception
            TraceLog("Error deleteDevice: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
        End Try
    End Sub
End Class
