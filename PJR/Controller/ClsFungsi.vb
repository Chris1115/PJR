Imports MySql.Data.MySqlClient

Public Class ClsFungsi
    Private utility As New Utility

    Public Shared Function SimpanQty_Bazar(ByVal tabel As String, ByVal prdcd As String, ByVal tgl_exp As String, ByVal qty As String) As ClsBazar
        Dim scon As MySqlConnection = ClsConnection.GetConnection.Clone
        Dim scmd As New MySqlCommand("", scon)
        Dim sdap As MySqlDataAdapter
        Dim dt As New DataTable
        Dim periode As String
        Dim result As New ClsBazar
        Dim recid As String = ""

        Try
            If scon.State = ConnectionState.Closed Then
                scon.Open()
            End If
            periode = Format(Date.Now, "yyMMdd")
            periode = periode.Substring(0, 2) + periode.Substring(2, 2)

            scmd.CommandText = "SELECT * from " & tabel & " where prdcd='" & prdcd & "' OR barcode='" & prdcd & "';"

            Console.WriteLine(scmd.CommandText)
            sdap = New MySqlDataAdapter(scmd)
            sdap.Fill(dt)

            For i As Integer = 0 To dt.Rows.Count - 1
                If tgl_exp = dt.Rows(i)("bulan_exp").ToString And (dt.Rows(i)("RECID").ToString = "P" Or dt.Rows(i)("RECID").ToString = "") Then

                    'tambah kolom terkait Memo 1893/CPS/22
                    scmd.CommandText = "UPDATE " & tabel & " SET FREQ = FREQ + 1  "
                    scmd.CommandText &= " WHERE (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') AND "
                    scmd.CommandText &= " bulan_exp = '" & tgl_exp & "' AND FREQ = 0"
                    scmd.ExecuteNonQuery()
                    'tambah kolom terkait Memo 1893/CPS/22
                    scmd.CommandText = "UPDATE " & tabel & " SET FREQ = FREQ + 1"
                    scmd.CommandText &= " WHERE (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') AND "
                    scmd.CommandText &= " bulan_exp = '" & tgl_exp & "'"
                    scmd.CommandText &= " AND DATE(UPDTIME) < CURDATE() AND FREQ <> 0 AND DATE(ADDTIME) <> CURDATE()"
                    scmd.ExecuteNonQuery()
                    IDM.Fungsi.TraceLog("log bazar SimpanQTY_bazar_P_2 : " & scmd.CommandText)

                    recid = "P"
                    scmd.CommandText = "UPDATE " & tabel & " SET"
                    scmd.CommandText &= " ttl = '" & qty & "',"
                    scmd.CommandText &= " RECID = '" & recid & "'"
                    'scmd.CommandText &= " UPDTIME = CURDATE()"
                    scmd.CommandText &= " WHERE (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') AND "
                    scmd.CommandText &= " bulan_exp = '" & tgl_exp & "'"

                    Console.WriteLine(scmd.CommandText)
                    IDM.Fungsi.TraceLog("log bazar SimpanQTY_bazar_P : " & scmd.CommandText)

                    scmd.ExecuteNonQuery()

                    GoTo next_step
                End If
            Next
            'tambah kolom terkait Memo 1893/CPS/22
            scmd.CommandText = "UPDATE " & tabel & " SET FREQ = FREQ + 1  "
            scmd.CommandText &= " WHERE (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') AND "
            scmd.CommandText &= " bulan_exp = '" & tgl_exp & "' AND FREQ = 0 "
            scmd.ExecuteNonQuery()
            'tambah kolom terkait Memo 1893/CPS/22
            scmd.CommandText = "UPDATE " & tabel & " SET FREQ = FREQ + 1, UPDTIME = CURDATE()"
            scmd.CommandText &= " WHERE (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') AND "
            scmd.CommandText &= " bulan_exp = '" & tgl_exp & "'"
            scmd.CommandText &= " AND DATE(UPDTIME) < CURDATE() AND FREQ <> 0  AND DATE(ADDTIME) <> CURDATE()"
            scmd.ExecuteNonQuery()

            recid = "B"
            scmd.CommandText = "INSERT IGNORE INTO " & tabel & " VALUES ("
            scmd.CommandText &= "'" & recid & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("tiperak").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("norak").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("noshelf").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("kirikanan").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("prdcd").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("singkat").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("barcode").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("com").ToString & "',"
            scmd.CommandText &= "'" & tgl_exp & "', "
            scmd.CommandText &= "'" & qty & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("SOID").ToString & "', "
            scmd.CommandText &= "'" & dt.Rows(0)("NO_PROP_BAZAR").ToString & "', "
            scmd.CommandText &= "CURDATE(), "
            scmd.CommandText &= "1, "
            scmd.CommandText &= "CURDATE(), "
            scmd.CommandText &= "'0001-01-01', "
            scmd.CommandText &= "'', "
            scmd.CommandText &= "'') "
            scmd.CommandText &= " ON DUPLICATE KEY UPDATE prdcd = values(prdcd), bulan_exp = values(bulan_exp),ttl=VALUES(ttl);"
            Console.WriteLine(scmd.CommandText)
            IDM.Fungsi.TraceLog("log bazar SimpanQTY_bazar_B : " & scmd.CommandText)

            scmd.ExecuteNonQuery()

next_step:
            result.BZR = New ClsBazar
            result.StatusExp = "2"
            result.BZR.Feedback = "Selesai update data"

            scmd.CommandText = "Select * from " & tabel & " where (prdcd = '" & prdcd & "' OR BARCODE = '" & prdcd & "') and bulan_exp = '" & tgl_exp & "'"
            sdap = New MySqlDataAdapter(scmd)
            sdap.Fill(dt)

            result.BZR.PRDCD = dt.Rows(0)("prdcd").ToString
            result.BZR.Deskripsi = dt.Rows(0)("singkat").ToString
            result.BZR.Tgl_exp = dt.Rows(0)("bulan_exp").ToString
            result.BZR.QTYTotal = dt.Rows(0)("ttl").ToString
        Catch ex As Exception
            IDM.Fungsi.TraceLog("Gagal update table SZ " & ex.Message & ex.StackTrace)
        Finally
            scon.Close()
        End Try

        Return result
    End Function
End Class
