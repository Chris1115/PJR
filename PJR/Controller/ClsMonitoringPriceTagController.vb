Imports MySql.Data.MySqlClient
Imports IDM.Fungsi

Public Class ClsMonitoringPriceTagController
    Private utility As New Utility

    Public Function CekTableMonitoringPriceTag_wdcp(ByVal tablename As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Mda As New MySqlDataAdapter("", Conn)

        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Mcom.CommandText = "CREATE TABLE IF NOT EXISTS `pos`.`" & tablename & "` ( "
            Mcom.CommandText &= "`modis` VARCHAR(100) DEFAULT NULL, "
            Mcom.CommandText &= "`PLU` VARCHAR(8) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`bisa_input` VARCHAR(5) NOT NULL, "
            Mcom.CommandText &= "`tanggal` DATE DEFAULT NULL, "
            Mcom.CommandText &= "`tanggalscan` DATETIME DEFAULT NULL, "
            Mcom.CommandText &= "`alasan_tidakscan` VARCHAR(200) DEFAULT NULL, "
            Mcom.CommandText &= "`NIK_Pramu` VARCHAR(20) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`Nama_Pramu` VARCHAR(50) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`NIK_PS` VARCHAR(20) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`Nama_PS` VARCHAR(50) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "`status` VARCHAR(200) NOT NULL DEFAULT '', "
            Mcom.CommandText &= "PRIMARY KEY (`modis`,`PLU`,`tanggal`), "
            Mcom.CommandText &= ") ENGINE=INNODB DEFAULT CHARSET=latin1;"
            Mcom.ExecuteNonQuery()

        Catch ex As Exception
            utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "CekTablePlano", Conn)
        Finally
            Conn.Close()
        End Try

        Return True
    End Function

    'Update MEMO 296/CPS/23 by Kukuh
    Public Function GetDeskripsiMonitoringPriceTag(ByVal tabel_name As String, ByVal barcode_plu As String, ByVal keterangan As String) As ClsMonitoringPriceTag
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim Result As New ClsMonitoringPriceTag
        Dim dt As New DataTable

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiMonitoringPriceTag (connection Nothing) " & vbCrLf & "PLU:" & barcode_plu)
            Return Result
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If

                Mcom.CommandText = "SELECT DISTINCT mpc.PLU plu, mpc.bisa_input bisa_input, mpc.tanggalscan tanggalscan "
                Mcom.CommandText &= "FROM monitoring_wdcp_ptag mpc LEFT JOIN barcode b "
                Mcom.CommandText &= "ON b.PLU = mpc.PLU WHERE (b.BARCD = '" & barcode_plu & "' OR mpc.PLU = '" & barcode_plu & "') "
                Mcom.CommandText &= "AND mpc.tanggal=DATE(NOW());"

                TraceLog(Mcom.CommandText)

                Dim sDap As New MySqlDataAdapter(Mcom)
                sDap.Fill(dt)

                TraceLog("MPC: Jumlah Data " & barcode_plu & ": " & dt.Rows.Count)

                If dt.Rows.Count > 0 Then
                    Result.barcode = dt.Rows(0).Item("plu")
                    Dim flagBisaInput = dt.Rows(0).Item("bisa_input")

                    TraceLog("MPC flag bisa_input: " & flagBisaInput)

                    If (flagBisaInput = "N" Or Trim(flagBisaInput) = "") And keterangan = "B" Then
                        Result.barcode = ""
                        Result.keterangan = "Plu hanya boleh di  scan, Silakan scan  barcode" 'Data tidak ada di tabel monitoring_wdcp_ptag
                        Result.setMenu = "3"
                    Else
                        TraceLog("Cek tgl scan " & barcode_plu & ": " & dt.Rows(0).Item("tanggalscan").ToString)
                        If IsDBNull(dt.Rows(0).Item("tanggalscan")) Then
                            Result.keterangan = "Item Ditemukan" 'Data ada dan belum di scan
                            Result.setMenu = "2"
                        Else
                            Result.barcode = ""
                            Result.keterangan = "Item Sudah di Scan" 'Data ada dan sudah di scan
                            Result.setMenu = "1"
                        End If
                    End If
                Else
                    Result.barcode = ""
                    Result.keterangan = "Item Tidak Ditemukan" 'Data tidak ada di tabel monitoring_wdcp_ptag
                    Result.setMenu = "1"
                End If

            Catch ex As Exception
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiMonitoringPriceTag", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiMonitoringPriceTag " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try
        End SyncLock

        Return Result
    End Function

    Public Function UpdateDataMPC(ByVal tabel_name As String, ByVal barcode_plu As String) As Boolean
        Dim Conn As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim Mcom As New MySqlCommand("", Conn)
        Dim dt As New DataTable

        If Conn Is Nothing Then
            utility.TraceLogTxt("Error - GetDeskripsiMonitoringPriceTag (connection Nothing)Then " & vbCrLf & "PLU:" & barcode_plu)
            Return False
            Exit Function
        End If

        SyncLock Conn
            Try
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If
                Mcom.CommandText = "UPDATE monitoring_wdcp_ptag SET tanggalscan = NOW() WHERE PLU = '" & barcode_plu & "';"
                Mcom.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Return False
                utility.Tracelog("Error", ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source, "GetDeskripsiMonitoringPriceTag", Conn)
                utility.TraceLogTxt("Error - GetDeskripsiMonitoringPriceTag " & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & ex.Source)
            Finally
                Conn.Close()
            End Try

        End SyncLock

    End Function

End Class
